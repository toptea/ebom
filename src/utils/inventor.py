"""
----------------
Inventor COM API
----------------
"""

from pathlib import Path
from utils.system import find_export_path, start_inventor
import win32com.client
import numpy as np
import pandas as pd
import textwrap


def application(silent=True, visible=True):
    """Inventor Application COM Object

    Start COM client session with Inventor, and create object 'mod' that will
    point to the Python COM wrapper for Inventor's type library. Recast 'app'
    as an instance of the Application class in the wrapper.

    Parameters
    ----------
    silent : bool
        controls whether an operation will proceed without prompting
    visible : bool
        sets the visibility of this application

    Returns
    -------
    obj
        Inventor Application COM Object
    """
    mod = win32com.client.gencache.EnsureModule(
        '{D98A091D-3A0F-4C3E-B36E-61F62068D488}', 0, 1, 0)
    app = win32com.client.Dispatch('Inventor.Application')
    try:
        app = mod.Application(app)
    except:
        app = mod.Application.Application(app)
    app.SilentOperation = silent
    app.Visible = visible
    return app


class Document:
    """Document

    The Document base class contains methods and properties for Inventor's
    Document COM object. Used to query inventor current file.

    Parameters
    ----------
    path : obj
        Path object from python pathlib module
    app : obj
        Inventor Application COM Object
    export_dir : str
        Export directory location

    Attributes
    ----------
    path : obj
        Path Object from python pathlib module
    app : obj
        Inventor Application COM Object
    doc : obj
        Inventor Document COM Object
    export_dir : str
        Export directory location
    """

    def __init__(self, path, app, export_dir=find_export_path()):
        self.app = app
        self.doc = self._load_document(path, app)
        self.export_dir = export_dir
        self.path = path

    @property
    def partcode(self):
        """str: return the file's partcode from it's full path"""
        return self.path.stem

    @staticmethod
    def _load_document(path, app):
        """Inventor Document Object

        Open the specified Inventor document.
        Check document type and bind the document COM object to the associate
        class in the wrapper.

        Parameters
        ----------
        path : obj
            Path Object from python pathlib module
        app : obj
            Inventor Application COM Object

        Returns
        -------
        obj
            Inventor Document COM Object
        """
        start_inventor()
        document_type_enum = {
            12289: 'UnnownDocument',
            12290: 'PartDocument',
            12291: 'AssemblyDocument',
            12292: 'DrawingDocument',
            12293: 'PresentationDocument',
            12294: 'DesignElementDocument',
            12295: 'ForeignModelDocument',
            12296: 'SATFileDocument',
            12297: 'NoDocument',
        }
        try:
            app.Documents.Open(str(path))
            document_type = document_type_enum[app.ActiveDocumentType]
            doc = win32com.client.CastTo(app.ActiveDocument, document_type)
            return doc
        except IOError as e:
            print(e + 'Unable to load Inventor document')

    @classmethod
    def via_active_document(cls, app, export_dir=find_export_path()):
        """Initialize class using model/drawing that was already open in Inventor"""
        try:
            path = Path(app.ActiveDocument.FullFileName)
            return cls(path, app, export_dir)
        except AttributeError:
            err = 'The file is not active in Inventor'
            raise IOError(err)

    def get_iproperties_data(self):
        """return a dictionary of useful iproperties of the model/drawing"""
        i = self.doc.PropertySets.Item("Inventor User Defined Properties")

        try:
            p_partcode = str(i.Item('Dwg_No')).strip()
        except:
            p_partcode = ''

        try:
            p_desc = str(i.Item('Component')).strip()
        except:
            p_desc = ''

        try:
            p_rev = int(i.Item('Revision'))
        except:
            p_rev = 1

        try:
            p_material = str(i.Item('Material')).strip()
        except:
            p_material = ''

        try:
            p_finish = str(i.Item('Finish')).strip()
        except:
            p_finish = ''

        try:
            p_drawn_by = str(i.Item('Drawn_by')).strip()
        except:
            p_drawn_by = ''

        try:
            p_drawn_dt = str(i.Item('Drawn_dt')).strip()
        except:
            p_drawn_dt = ''

        try:
            p_title = str(i.Item('Title')).strip()
        except:
            p_title = ''

        iprop = {
            'partcode': p_partcode,
            'desc': p_desc,
            'rev': p_rev,
            'material': p_material,
            'finish': p_finish,
            'drawn_by': p_drawn_by,
            'drawn_dt': p_drawn_dt,
            'title': p_title
        }
        return iprop

    def close(self):
        """Close Document

        Close current Inventor document without saving.
        """
        self.doc.Close(SkipSave=True)


class Drawing(Document):
    """Drawing Document

    The Drawing Class contains methods and properties for Inventor's
    DrawingDocument COM object. Used to query idw file.
    """

    def get_drawing_sheet_size(self):
        """Sheet Size

        Get the size of the sheet.
        """
        drawing_sheet_size_enum = {
            9993: 'A0', 9994: 'A1', 9995: 'A2', 9996: 'A3', 9997: 'A4'
        }
        return drawing_sheet_size_enum[self.doc.Sheets(1).Size]

    def extract_part_list(self, lvl=1):
        df = pd.DataFrame()

        sheets = self.doc.Sheets
        for nsheet in range(1, self.doc.Sheets.Count + 1):

            partlists = sheets(nsheet).PartsLists
            for n in range(1, partlists.Count + 1):
                rs = self._extract(partlists(n))
                df = df.append(rs)

        df['Assembly'] = self.get_iproperties_data()['partcode']
        df['Assembly_Name'] = self.get_iproperties_data()['desc']
        df['LVL'] = lvl
        try:
            df = df[['Assembly', 'Assembly_Name', 'LVL', 'ITEM', 'QTY', 'Dwg_No', 'Component']]
            df['ITEM'] = df['ITEM'].astype(int)
            df['QTY'] = df['QTY'].astype(int)
            df['Assembly_Name'] = df['Assembly_Name'].str.strip()
            df['Component'] = df['Component'].str.strip()
        except KeyError:
            err = textwrap.dedent(
                """
                Unable to extract part list.
                ---------------------------------------------------------------
                Make sure the assembly drawing have a part list.
                Part list should also have 'ITEM', 'QTY', 'Dwg_No' and 'Component' columns.
                """
            )
            raise IOError(err)
        return df

    @staticmethod
    def _extract(part_list):
        columns = part_list.PartsListColumns
        rows = part_list.PartsListRows

        data = {}
        for ncol in range(1, columns.Count + 1):

            column_data = []
            for nrow in range(1, rows.Count + 1):
                try:
                    if rows.Item(nrow).Visible:
                        cell = str(rows.Item(nrow).Item(ncol)).strip()
                    else:
                        cell = None
                except:
                    cell = None

                if cell is not None:
                    column_data.append(cell)

            title = columns(ncol).Title
            data[title] = column_data

        df = pd.DataFrame(data)
        return df


class Assembly(Document):
    """Assembly Document

    The Assembly Class contains methods and properties for Inventor's
    AssemblyDocument COM object. Used to query iam file.
    """

    def export_bom(self):
        """Assembly BOM

        Export the assembly's bom to an excel spreadsheet.
        """
        path = self.export_dir.joinpath(self.partcode).joinpath('bom.xlsx')
        bom = self.doc.ComponentDefinition.BOM
        bom.StructuredViewFirstLevelOnly = False
        bom.StructuredViewEnabled = True
        bom.BOMViews.Item("Structured").Export(path, 74498)


class Part(Document):
    """Part Document

    The Part Class contains methods and properties for Inventor's
    PartDocument COM object. Used to query ipt file.
    """

    def is_blank_file(self):
        """
        return true if there's no features in the document.
        use to prevent further exceptions.
        """
        return self.doc.ComponentDefinition.Features.Count == 0

    def is_sheet_fabrication(self):
        """
        return true if it's a sheet fabrication part.
        Can only check by using the subtype GUID.
        """
        return self.doc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"

    def is_turn_part(self):
        """
        return true if the first feature is an extrude and the first
        sketch contains a circle, or return true if the first feature
        is a revolution feature.
        """
        c = self.doc.ComponentDefinition
        if c.Features.Count == 0 or c.Sketches.Count == 0:
            return False
        else:
            has_extrude = c.Features(1).Type == 83910656
            has_revolve = c.Features(1).Type == 83914240
            has_sketch_circle = c.Sketches(1).SketchCircles.Count
            return (has_extrude and has_sketch_circle) or has_revolve

    def check_ipt_subtype(self):
        if self.is_blank_file():
            return 'Blank'
        elif self.is_sheet_fabrication():
            return 'Sheet'
        elif self.is_turn_part():
            return 'Turning'
        else:
            return 'Machining'

    def pull_mass_properties_data(self):
        m = self.doc.ComponentDefinition.MassProperties
        m.Accuracy = 24579  # High
        mprop = {
            'iprop_volume': round(m.Volume, 3),
            'iprop_area': round(m.Area, 3),
            # 'p_mass' = round(m.Mass, 3)
        }
        return mprop

    def pull_bounding_box_data(self):
        rangebox = self.doc.ComponentDefinition.RangeBox
        x_min = rangebox.MinPoint.X
        y_min = rangebox.MinPoint.Y
        z_min = rangebox.MinPoint.Z
        x_max = rangebox.MaxPoint.X
        y_max = rangebox.MaxPoint.Y
        z_max = rangebox.MaxPoint.Z

        x = round(x_max - x_min, 3)
        y = round(y_max - y_min, 3)
        z = round(z_max - z_min, 3)

        bbox = {
            'bbox_min': min(x, y, z),
            'bbox_med': np.median([x, y, z]),
            'bbox_max': max(x, y, z),
            'bbox_area': 2 * x * y + 2 * x * z + 2 * y * z,
            'bbox_vol': x * y * z,
        }
        return bbox

    def pull_cad_feature_data(self):

        f = self.doc.ComponentDefinition.Features
        feature = {
            'feat_total': f.Count,
            'feat_alias_freeform': f.AliasFreeformFeatures.Count,
            'feat_bend_part': f.BendPartFeatures.Count,
            'feat_boss': f.BossFeatures.Count,
            'feat_boundary_patch': f.BoundaryPatchFeatures.Count,
            'feat_chamfer': f.ChamferFeatures.Count,
            'feat_circular_pattern': f.CircularPatternFeatures.Count,
            'feat_client': f.ClientFeatures.Count,
            'feat_coil': f.CoilFeatures.Count,
            'feat_combine': f.CombineFeatures.Count,
            'feat_decal': f.DecalFeatures.Count,
            'feat_delete_face': f.DeleteFaceFeatures.Count,
            'feat_direct_edit': f.DirectEditFeatures.Count,
            'feat_emboss': f.EmbossFeatures.Count,
            'feat_extend': f.ExtendFeatures.Count,
            'feat_extrude': f.ExtrudeFeatures.Count,
            'feat_face_draft': f.FaceDraftFeatures.Count,
            'feat_fillet': f.FilletFeatures.Count,
            'feat_freeform': f.FreeformFeatures.Count,
            'feat_grill': f.GrillFeatures.Count,
            'feat_hole': f.HoleFeatures.Count,
            'feat_i': f.iFeatures.Count,
            'feat_knit': f.KnitFeatures.Count,
            'feat_lip': f.LipFeatures.Count,
            'feat_loft': f.LoftFeatures.Count,
            'feat_mirror': f.MirrorFeatures.Count,
            'feat_move_face': f.MoveFaceFeatures.Count,
            'feat_move': f.MoveFeatures.Count,
            'feat_nonparametric_base': f.NonParametricBaseFeatures.Count,
            'feat_rectangular_pattern': f.RectangularPatternFeatures.Count,
            'feat_reference': f.ReferenceFeatures.Count,
            'feat_replace_face': f.ReplaceFaceFeatures.Count,
            'feat_rest': f.RestFeatures.Count,
            'feat_revolve': f.RevolveFeatures.Count,
            'feat_rib': f.RibFeatures.Count,
            'feat_ruled_surface': f.RuledSurfaceFeatures.Count,
            'feat_rule_fillet': f.RuleFilletFeatures.Count,
            'feat_sculpt': f.SculptFeatures.Count,
            'feat_shell': f.ShellFeatures.Count,
            'feat_snap_fit': f.SnapFitFeatures.Count,
            'feat_split': f.SplitFeatures.Count,
            'feat_sweep': f.SweepFeatures.Count,
            'feat_thicken': f.ThickenFeatures.Count,
            'feat_thread': f.ThreadFeatures.Count,
            'feat_trim': f.TrimFeatures.Count,
        }
        return feature

    def pull_brep_data(self):
        bodies = self.doc.ComponentDefinition.SurfaceBodies

        brep = {
            'brep_body': 0,
            'brep_face': 0,
            'brep_edgeloop': 0,
            'brep_edge': 0,
        }

        brep['brep_body'] += bodies.Count
        for body in bodies:
            brep['brep_face'] += body.Faces.Count
            for face in body.Faces:
                brep['brep_edgeloop'] += face.EdgeLoops.Count
                for edgeloop in face.EdgeLoops:
                    brep['brep_edge'] += edgeloop.Edges.Count

        return brep

    def pull_all(self):
        iprop = self.pull_iproperties_data()
        mprop = self.pull_mass_properties_data()
        bbox = self.pull_bounding_box_data()
        feature = self.pull_cad_feature_data()
        brep = self.pull_brep_data()
        result = {**iprop, **mprop, **brep, **feature, **bbox}
        return result

