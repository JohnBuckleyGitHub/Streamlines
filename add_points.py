import pythoncom
import win32com.client
import csv
import math


class PointAdder(object):

    def __init__(self):
        try:
            self.catia = win32com.client.Dispatch("CATIA.Application")
            self.doc = self.catia.ActiveDocument
            self.part = self.doc.Part
            self.hyfac = self.part.HybridShapeFactory
            self.sel = self.doc.Selection
        except AttributeError:
            print('CATIA must be loaded and a part activated')
        else:
            self.prelim_gset_list = []
            print("CATPart '" + str(self.part.Name) + "' selected")
            # gname = self.add_points('NN')
            # self.make_curve(gname)
            # jname = self.join_lines(gname, 'j1')
            # cname = self.create_ccurve(jname, 'c1')
            # self.delete_all_but_name(gname, cname)

    def add_file_locations(self):
        self.data_folder = ('C:\\PME_Mirror\\GM_IndyCar\\Vehicle_Data\\Aero\\CFD\\RESULTS\\PHASE09-RC_DEV\\' +
                            'GMICS-P09-C178 Data')
        self.pv_folder = 'C:\\Program Files\\ParaView 5.0.0\\bin'
        self.dat_file_folder = 'C:\\Users\\Johns Lenovo\\Documents\\GitHub\\Streamlines'
        self.dat_file = 'group_01_1.dat'

    def test(self):
        sline_hbody = self.part.HybridBodies.Add()
        sline_hbody_ref = self.part.CreateReferenceFromObject(sline_hbody)
        self.part.HybridShapeFactory.ChangeFeatureName(sline_hbody_ref, "Streamlines")
        seed_pt = self.part.HybridBodies.Item('Geometrical Set.1').HybridShapes.Item('Point.181')
        sline = self.Streamline(self, seed_pt)
        sline.open_dat_file('group_01_1.dat')
        sline.create_lines()
        sline.create_join()
        cc = sline.create_ccurve()
        sline_hbody.AppendHybridShape(cc)
        self.sel.Clear()
        self.sel.Add(sline.pre_hbody)
        self.sel.Delete()

    def add_points(self):
        hbody = self.part.HybridBodies.Add()
        p = math.pi/180
        for i in range(180):
            point = self.hyfac.AddNewPointCoord(i, 100*math.sin(i*p), 100*math.cos(i*p))
            point.Name = 'Point_' + str(i+1)
            hbody.AppendHybridShape(point)
        self.part.update()
        self.prelim_gset_list.append(hbody)

    def make_curve(self, g_set_name):
        hbody = self.part.HybridBodies.Item(g_set_name)
        self.line_ref_list = []
        for i in range(1, hbody.HybridShapes.Count + 1 - 1):
            pt_1 = hbody.HybridShapes.Item(i)
            ref_1 = self.part.CreateReferenceFromObject(pt_1)
            pt_2 = hbody.HybridShapes.Item(i+1)
            ref_2 = self.part.CreateReferenceFromObject(pt_2)
            line = self.hyfac.AddNewLinePtPt(ref_1, ref_2)
            self.line_ref_list.append(self.part.CreateReferenceFromObject(line))
            hbody.AppendHybridShape(line)
            # print(pt_1.Name)
        self.part.update()

    def join_lines(self, g_set_name, join_name):
        hbody = self.part.HybridBodies.Item(g_set_name)
        join_created = 0
        for line in self.line_ref_list:
            if join_created == 0:
                line1 = line
                join_created = 1
            elif join_created == 1:
                join = self.hyfac.AddNewJoin(line1, line)
                join_created = 2
            else:
                join.AddElement(line)
        self.set_join_params(join)
        if self.check_for_feature(g_set_name, join_name):
            join_name = str(join_name) + "_1"
        join.Name = join_name
        hbody.AppendHybridShape(join)
        self.part.update()
        print(join_name)
        return join_name

    def create_ccurve(self, join_name, curve_name):
        hbody = self.get_hbody(join_name)
        join = hbody.HybridShapes.Item(join_name)
        ref1 = self.part.CreateReferenceFromObject(join)
        cc = self.hyfac.AddNewCurveSmooth(ref1)
        cc.SetTangencyThreshold(0.5)
        cc.CurvatureThresholdActivity = False
        cc.MaximumDeviationActivity = True
        cc.SetMaximumDeviation(0.1)
        cc.TopologySimplificationActivity = True
        cc.CorrectionMode = 3
        hbody.AppendHybridShape(cc)
        self.part.update()
        ref2 = self.part.CreateReferenceFromObject(cc)
        cc2 = self.hyfac.AddNewCurveDatum(ref2)
        if self.check_for_feature(hbody, curve_name):
            curve_name = str(curve_name) + "_1"
        cc2.Name = curve_name
        hbody.AppendHybridShape(cc2)
        self.part.update()
        self.hyfac.DeleteObjectForDatum(ref2)
        return curve_name

    def delete_all_but_name(self, hbody, feature_name):
        if isinstance(hbody, str):
            hbody = self.part.HybridBodies.Item(hbody)
        sel = self.doc.Selection
        sel.Clear()
        for i in range(hbody.HybridShapes.Count):
            feature = hbody.HybridShapes.Item(i+1)
            if feature.Name != feature_name:
                sel.Add(feature)
        sel.Delete()

    def get_hbody(self, feature_name):
        for i in range(self.part.HybridBodies.Count):
            try:
                hbody = self.part.HybridBodies.Item(i+1)
                feature = hbody.HybridShapes.Item(feature_name)
            except pythoncom.com_error:
                pass
            else:
                return hbody

    def check_for_hbody(self, hbody_name):
        try:
            hbody = self.part.HybridBodies.Item(hbody_name)
        except pythoncom.com_error:
            return False
        else:
            return True

    def check_for_feature(self, hbody_name, feature_name):
        try:
            hbody = self.part.HybridBodies.Item(hbody_name)
            feature = hbody.HybridShapes.Item(feature_name)
        except pythoncom.com_error:
            return False
        else:
            return True

    class Streamline(object):

        def __init__(self, parent, seed_pt_obj):
            self.part = parent.part
            self.hyfac = parent.hyfac
            self.seed_pt_obj = seed_pt_obj
            self.seed_name = seed_pt_obj.Name
            self.pre_hbody = self.part.HybridBodies.Add()
            self.final_hbody = None
            self.ref_pt_list = []
            self.ref_line_list = []
            
            self.join_ref = None

        def open_dat_file(self, dat_file):
            with open(dat_file, 'r') as f:
                dat_reader = csv.reader(f, delimiter=' ')
                for row in dat_reader:
                    self.add_point(row[0], row[1:])
            self.part.update()

        def add_point(self, name, coord_tuple):
            point = self.hyfac.AddNewPointCoord(coord_tuple[0], coord_tuple[1], coord_tuple[2])
            point.Name = name
            self.pre_hbody.AppendHybridShape(point)
            self.ref_pt_list.append(self.part.CreateReferenceFromObject(point))

        def create_lines(self):
            prev_ref = None
            for ref in self.ref_pt_list:
                if prev_ref:
                    line = self.hyfac.AddNewLinePtPt(prev_ref, ref)
                    self.ref_line_list.append(self.part.CreateReferenceFromObject(line))
                    self.pre_hbody.AppendHybridShape(line)
                prev_ref = ref
            self.part.update()

        def create_join(self):
            join_created = 0
            for line in self.ref_line_list:
                if join_created == 0:
                    line1 = line
                    join_created = 1
                elif join_created == 1:
                    join = self.hyfac.AddNewJoin(line1, line)
                    join_created = 2
                else:
                    join.AddElement(line)
            set_join_params(join)
            self.pre_hbody.AppendHybridShape(join)
            self.join_ref = self.part.CreateReferenceFromObject(join)
            self.part.update()

        def create_ccurve(self):
            cc = self.hyfac.AddNewCurveSmooth(self.join_ref)
            cc.SetTangencyThreshold(0.5)
            cc.CurvatureThresholdActivity = False
            cc.MaximumDeviationActivity = True
            cc.SetMaximumDeviation(0.1)
            cc.TopologySimplificationActivity = True
            cc.CorrectionMode = 3
            # self.final_hbody = self.part.HybridBodies.Add()
            self.pre_hbody.AppendHybridShape(cc)
            self.part.update()
            ref2 = self.part.CreateReferenceFromObject(cc)
            cc2 = self.hyfac.AddNewCurveDatum(ref2)
            cc2.Name = 'Streamline_' + str(self.seed_name)
            # hbody.AppendHybridShape(cc2)
            # self.part.update()
            self.hyfac.DeleteObjectForDatum(ref2)
            return cc2

    def ref_junk_code(self):
        part = self.catia.ActiveDocument.Part
        hbody = part.HybridBodies.Add()
        referencebody = part.CreateReferenceFromObject(hbody)
        part.HybridShapeFactory.ChangeFeatureName(referencebody, "New Name")
        point = part.HybridShapeFactory.AddNewPointCoord(10, 20, 30)
        hbody.AppendHybridShape(point)
        part.update()
        nn = part.HybridBodies.Item("New Name")
        np = nn.HybridShapes.Item("Point.1")
        np.Name = 'anything'
        nn.HybridShapes.Item(1).Name = 'licky'
        part.Parameters.Item("Part1\\New Name\\licky\\X").Value = 100
        number_of_items = nn.HybridShapes.Count
        referencebody = self.part.CreateReferenceFromObject(hbody)
        g_set_name = 'licky'
        if self.check_for_hbody(g_set_name):
            g_set_name = str(g_set_name) + "_1"
        self.hyfac.ChangeFeatureName(referencebody, g_set_name)


def set_join_params(join):
    join.SetConnex(1)
    join.SetManifold(1)
    join.SetSimplify(0)
    join.SetSuppressMode(0)
    join.SetDeviation(0.001)
    join.SetAngularToleranceMode(0)
    join.SetAngularTolerance(0.5)
    join.SetFederationPropagation(0)

