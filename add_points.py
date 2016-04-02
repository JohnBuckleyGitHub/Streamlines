import pythoncom
import win32com.client
import math


class PointAdder(object):

    def __init__(self):
        self.catia = win32com.client.Dispatch("CATIA.Application")
        self.doc = self.catia.ActiveDocument
        self.part = self.doc.Part
        self.hyfac = self.part.HybridShapeFactory
        self.g_set_list = []
        gname = self.add_points('NN')
        self.make_curve(gname)
        jname = self.join_lines(gname, 'j1')
        cname = self.create_ccurve(jname, 'c1')
        self.delete_all_but_name(gname, cname)

    def add_points(self, g_set_name):
        self.part = self.catia.ActiveDocument.Part
        self.hyfac = self.part.HybridShapeFactory
        hbody = self.part.HybridBodies.Add()
        referencebody = self.part.CreateReferenceFromObject(hbody)
        if self.check_for_hbody(g_set_name):
            g_set_name = str(g_set_name) + "_1"
        self.hyfac.ChangeFeatureName(referencebody, g_set_name)
        p = math.pi/180
        for i in range(180):
            point = self.hyfac.AddNewPointCoord(i, 100*math.sin(i*p), 100*math.cos(i*p))
            point.Name = 'Point_' + str(i+1)
            hbody.AppendHybridShape(point)
        self.part.update()
        print(g_set_name)
        return g_set_name

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









    def set_join_params(self, join):
        join.SetConnex(1)
        join.SetManifold(1)
        join.SetSimplify(0)
        join.SetSuppressMode(0)
        join.SetDeviation(0.001)
        join.SetAngularToleranceMode(0)
        join.SetAngularTolerance(0.5)
        join.SetFederationPropagation(0)



    def ref_junk_code(self):
        part = self.catia.ActiveDocument.Part
        myHBody = part.HybridBodies.Add()
        referencebody = part.CreateReferenceFromObject(myHBody)
        part.HybridShapeFactory.ChangeFeatureName(referencebody, "New Name")
        point = part.HybridShapeFactory.AddNewPointCoord(10, 20, 30)
        myHBody.AppendHybridShape(point)
        part.update()
        nn = part.HybridBodies.Item("New Name")
        np = nn.HybridShapes.Item("Point.1")
        np.Name = 'anything'
        nn.HybridShapes.Item(1).Name = 'licky'
        part.Parameters.Item("Part1\\New Name\\licky\\X").Value = 100
        number_of_items = nn.HybridShapes.Count


