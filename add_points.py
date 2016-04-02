import pythoncom
import win32com.client
import csv
import subprocess
import numpy as np


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
            self.add_file_locations()

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
        seed_pt = self.part.HybridBodies.Item('Geometrical Set.1').HybridShapes.Item('Point.1')
        sline = self.Streamline(self, seed_pt)
        sline.open_dat_file('Point.2_1.dat')
        sline.create_lines()
        sline.create_join()
        cc = sline.create_ccurve()
        sline_hbody.AppendHybridShape(cc)
        self.delete_with_selection(sline.pre_hbody)

    def delete_with_selection(self, obj):
        self.sel.Clear()
        self.sel.Add(obj)
        self.sel.Delete()

    def point_list_from_selection(self):
        self.sline_list = []
        for i in range(self.sel.Count):
            if self.sel.Item(i+1).Type == 'HybridShapePointCoord':
                hspoint = self.Point(self.sel.Item(i+1))
                self.sline_list.append(hspoint)

    def create_dat_file(self):
        header = ['group_name', 'x_location', 'y_location', 'z_location', 'direction(1-forward,2-backward)']
        with open('pvstream_seed.dat', 'w', newline='') as f:
            dat_writer = csv.writer(f, delimiter=' ')
            dat_writer.writerow(header)
            for pt in self.sline_list:
                dat_writer.writerow(pt.row)

    def run_script(self):
        pvpy = '"' + self.pv_folder + '\\' + 'pvpython.exe"'
        pvscript = 'pvstreamextract2.py'
        xmf = '"' + self.data_folder + '\\' + 'xdmf_avg0006000.xmf"'
        pv_seed = 'pvstream_seed.dat'
        geom_info = ' 26 30 121'
        arg_string = pvpy + ' ' + pvscript + ' ' + xmf + ' ' + pv_seed + geom_info
        subprocess.call(arg_string, shell=True)

    class Point(object):

        def __init__(self, hspoint):
            self.name = hspoint.Value.Name
            self.x = hspoint.Value.X.Value
            self.y = hspoint.Value.Y.Value
            self.z = hspoint.Value.Z.Value
            self.xyz = (self.x, self.y, self.z)
            self.row = [self.name, self.x/1000, self.y/1000, self.z/1000, 1]

    class Streamline(object):

        def __init__(self, parent, seed_pt_obj):
            self.part = parent.part
            self.hyfac = parent.hyfac
            self.seed_pt_obj = seed_pt_obj
            self.seed_name = seed_pt_obj.Name
            self.pre_hbody = self.part.HybridBodies.Add()
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
            sc_tup = tuple(float(x)*1000 for x in coord_tuple)
            point = self.hyfac.AddNewPointCoord(sc_tup[0], sc_tup[1], sc_tup[2])
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
            base_dev = 0.1
            for i in range(30):
                try:
                    dev = base_dev * (i+1)
                    self.set_ccurve(dev)
                    # cc.SetTangencyThreshold(0.5)
                    # cc.CurvatureThresholdActivity = False
                    # cc.MaximumDeviationActivity = True
                    # cc.SetMaximumDeviation(0.1)
                    # cc.TopologySimplificationActivity = True
                    # cc.CorrectionMode = 3
                    self.pre_hbody.AppendHybridShape(cc)
                    self.part.update()
                except pythoncom:
                    print('Fail on ' + str(dev) + 'mm')
                    pass
                else:
                    print('Curve accuracy is:' + str(dev))
                    break
            ref2 = self.part.CreateReferenceFromObject(cc)
            cc2 = self.hyfac.AddNewCurveDatum(ref2)
            cc2.Name = 'Streamline_' + str(self.seed_name)
            self.hyfac.DeleteObjectForDatum(ref2)
            return cc2

        def set_ccurve(cc, dev):
            cc.SetTangencyThreshold(0.5)
            cc.CurvatureThresholdActivity = False
            cc.MaximumDeviationActivity = True
            cc.SetMaximumDeviation(dev)
            cc.TopologySimplificationActivity = True
            cc.CorrectionMode = 3

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