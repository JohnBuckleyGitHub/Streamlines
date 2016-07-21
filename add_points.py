import pythoncom
import win32com.client
import csv
import subprocess
import pdb
import time
import numpy as np
import h5py
from lxml import etree as ET


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
        # self.data_folder = ''
        self.data_folder = ('C:\\PME_Mirror\\GM_IndyCar\\Vehicle_Data\\Aero\\CFD\\RESULTS\\PHASE09-RC_DEV\\' +
                            'GMICS-P09-C184\\')
        self.pv_folder = 'C:\\Program Files\\ParaView 5.0.0\\bin'
        self.dat_file_folder = 'C:\\Users\\Johns Lenovo\\Documents\\GitHub\\Streamlines'
        self.dat_file = 'group_01_1.dat'

    def draw_strlns(self, dat_file=None, max_time=None):
        strln_hbody = self.create_hbody('Streamlines')
        if not dat_file:
            for strln_pt in self.strln_list:
                strln = self.Streamline(self, seed_pt_obj=strln_pt, max_time=max_time)
                strln.open_dat_file()
                self._draw_strln(strln, strln_hbody)
        else:
            strln = self.Streamline(self, max_time=max_time)
            strln.open_dat_file(dat_file)
            self._draw_strln(strln, strln_hbody)

    def _draw_strln(self, strln, hbody):
        strln.create_lines()
        strln.create_join()
        cc = strln.create_ccurve()
        if cc is not None:
            hbody.AppendHybridShape(cc)
        self.delete_with_selection(strln.pre_hbody)

    def delete_with_selection(self, obj):
        self.sel.Clear()
        self.sel.Add(obj)
        self.sel.Delete()

    def get_strlns_from_selection(self):
        self.start_time = time.time()
        self.point_list_from_selection()
        self.time1 = time.time()
        self.create_dat_file()
        self.time2 = time.time()
        self.run_script()
        self.time3 = time.time()
        self.draw_strlns(max_time=0.01)
        self.finish_time = time.time()

    def point_list_from_selection(self):
        self.strln_list = []
        for i in range(self.sel.Count):
            if self.sel.Item(i+1).Type == 'HybridShapePointCoord':
                hspoint = self.SeedPoint(self, self.sel.Item(i+1), direc=1)
                hspoint2 = self.SeedPoint(self, self.sel.Item(i+1), direc=2)
                self.strln_list.append(hspoint)
                self.strln_list.append(hspoint2)

    def create_dat_file(self):
        header = ['group_name', 'x_location', 'y_location', 'z_location', 'direction(1-forward,2-backward)']
        with open('pvstream_seed.dat', 'w', newline='') as f:
            dat_writer = csv.writer(f, delimiter=' ')
            dat_writer.writerow(header)
            for pt in self.strln_list:
                dat_writer.writerow(pt.row)

    def run_script(self):
        pvpy = '"' + self.pv_folder + '\\' + 'pvpython.exe"'
        pvscript = 'pvstreamextract2.py'
        xmf_inp = '"' + self.data_folder + 'xdmf_avg0006000_vel_only.xmf"'  # xdmf_avg0006000.xmf"'
        geom_info = ' 26 30 121'  # -1 -1 0 1 0 1.5'
        # self.trim_hd5(xmf_inp, geom_info)
        # xmf = '"' + self.data_folder + '\\' + 'region.xmf"'
        pv_seed = 'pvstream_seed.dat'
        arg_string = pvpy + ' ' + pvscript + ' ' + xmf_inp + ' ' + pv_seed + geom_info
        print(arg_string)
        subprocess.call(arg_string, shell=True)

    def trim_hd5(self, xmf_inp, allargs):
        xmin = np.array([allargs[0], allargs[1], allargs[2]])
        xmax = np.array([allargs[3], allargs[4], allargs[5]])
        region_file = 'region.xmf'
        write_region_xmf(xmf_inp, xmin, xmax, outputfile=region_file)

    def _check_for_hbody(self, hbody_name):
        try:
            hbody = self.part.HybridBodies.Item(hbody_name)
        except pythoncom.com_error:
            return None
        else:
            return hbody

    def create_hbody(self, hbody_name, duplicate=False):
        hbody = self._check_for_hbody(hbody_name)
        if not hbody:
            hbody = self.part.HybridBodies.Add()
            # hbody_ref = self.part.CreateReferenceFromObject(hbody)
            hbody.Name = hbody_name
        return hbody

    class SeedPoint(object):

        def __init__(self, parent, hspoint, direc=1):
            self.name = hspoint.Value.Name
            self.x = hspoint.Value.X.Value
            self.y = hspoint.Value.Y.Value
            self.z = hspoint.Value.Z.Value
            self.xyz = (self.x, self.y, self.z)
            self.direc = direc
            self.row = [self.name, self.x/1000, self.y/1000, self.z/1000, self.direc]
            self.hbody = hspoint.Value.Parent.Parent.Name
            self.ref = parent.part.HybridBodies.Item(self.hbody).HybridShapes.Item(self.name)

    class Streamline(object):

        def __init__(self, parent, seed_pt_obj=None, max_time=None):
            self.part = parent.part
            self.hyfac = parent.hyfac
            self.pre_hbody = self.part.HybridBodies.Add()
            self.ref_pt_list = []
            self.ref_line_list = []
            self.join_ref = None
            if not max_time:
                max_time = 999
            self.max_time = max_time
            self.seed_pt_obj = seed_pt_obj
            if seed_pt_obj:
                self.seed_name = seed_pt_obj.name + '_' + str(seed_pt_obj.direc)

        def open_dat_file(self, dat_file=None):
            if not dat_file:
                dat_file = self.seed_name + '.dat'
            self.dat_file = dat_file
            with open(self.dat_file, 'r') as f:
                dat_reader = csv.reader(f, delimiter=' ')
                for row in dat_reader:
                    if float(row[0]) > self.max_time:
                        break
                    elif -float(row[0]) > self.max_time:
                        continue
                    self.add_point(row[0], row[1:])
            self.part.update()

        def add_point(self, name, coord_tuple):
            sc_tup = tuple(float(x)*1000 for x in coord_tuple)
            point = self.hyfac.AddNewPointCoord(sc_tup[0], sc_tup[1], sc_tup[2])
            point.Name = "{0:.5f}".format(float(name))
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
            if join_created < 2:
                return
            set_join_params(join)
            self.pre_hbody.AppendHybridShape(join)
            self.join_ref = self.part.CreateReferenceFromObject(join)
            self.part.update()

        def create_ccurve(self):
            if self.join_ref is None:
                return
            created = False
            cc = self.hyfac.AddNewCurveSmooth(self.join_ref)
            base_dev = 0.1
            dev = base_dev
            try:
                self.set_ccurve(cc, dev)
                self.pre_hbody.AppendHybridShape(cc)
                self.part.update()
            except pythoncom.com_error:
                for i in range(50):
                    try:
                        print('Fail on ' + str(dev) + 'mm')
                        dev += base_dev
                        cc.SetMaximumDeviation(dev)
                        self.part.update()
                    except pythoncom.com_error:
                        pass
                    else:
                        created = True
                        break
            else:
                created = True
            if not created:
                return None
            print(str(self.seed_name) + ': Curve accuracy is:' + str(dev))
            ref2 = self.part.CreateReferenceFromObject(cc)
            cc2 = self.hyfac.AddNewCurveDatum(ref2)
            if not self.seed_name:
                self.seed_name = self.dat_file[:-3]
            cc2.Name = 'Streamline_' + str(self.seed_name)
            self.hyfac.DeleteObjectForDatum(ref2)
            return cc2

        def set_ccurve(self, cc, dev):
            cc.SetTangencyThreshold(0.5)
            cc.CurvatureThresholdActivity = False
            cc.MaximumDeviationActivity = True
            cc.SetMaximumDeviation(dev)
            cc.TopologySimplificationActivity = True
            cc.CorrectionMode = 3


def set_join_params(join):
    join.SetConnex(1)
    join.SetManifold(1)
    join.SetSimplify(0)
    join.SetSuppressMode(0)
    join.SetDeviation(0.001)
    join.SetAngularToleranceMode(0)
    join.SetAngularTolerance(0.5)
    join.SetFederationPropagation(0)


def write_region_xmf(xmffile, min, max, outputfile='region.xmf'):
    """Write out an xmf file describing a region of the domain.
    Any grid block that is outside of the bounding box will
    not be written to the new region xmf file.

    Parameters
    ----------
    xmffile : str
        Name of the input xmf file.
    min : list
        Bounding box minimums in the x-, y- and z-directions.
    max : list
        Bounding box maximums in the x-, y- and z-directions.
    outputfile : str, optional
        Output xmf file name.

    """

    tree = ET.parse(xmffile)
    root = tree.getroot()
    try:
        domain = root.find('.//Domain[@Name="Raven_Grid"]')
        volume_and_boundaries = root.findall('.//Domain[@Name="Raven_Grid"]/Grid')
        volume = root.find('.//Domain[@Name="Raven_Grid"]/Grid[@Name="Raven_Grid"]')
        grids = volume.findall('.//Grid')
        converted = False
    except AttributeError:
        domain = root.find('.//Domain[@Name="Raven Grid"]')
        volume_and_boundaries = root.findall('.//Domain[@Name="Raven Grid"]/Grid')
        volume = root.find('.//Domain[@Name="Raven Grid"]/Grid[@Name="Volume"]')
        grids = volume.findall('.//Grid')
        converted = True
    grid_files = {}
    include_grid = {}

    # Loop over all the grids in the xmf file.
    for n, grid in enumerate(grids):
        grid_name = grid.attrib['Name']
        include_grid[grid_name] = 1
        geometry = grid.find('.//Geometry/DataItem[@Name="Nodes"]')
        # Get the HDF5 file name and the dataset for the grid
        grid_file, dataset = geometry.text.split(':')[:2]
        if grid_file not in grid_files.keys():
            # If we haven't already opened this file, then we need to.
            grid_files[grid_file] = h5py.File(grid_file, 'r')
        # Read the xyz dataset and compute the range for this block.
        xyz = grid_files[grid_file][dataset][:]
        xyz_min = xyz.min(0)
        xyz_max = xyz.max(0)
        for i in xrange(3):
            if xyz_min[i] > max[i] or xyz_max[i] < min[i]:
                # If this grid is outside the bounding box
                # we do not need to keep it.
                include_grid[grid_name] = 0
                break

    for f in grid_files.values():
        f.close()

    # Remove the unnecessary grids from the xmf file.
    if converted:
        for grid in volume_and_boundaries:
            subgrids = grid.findall('.//Grid')
            for subgrid in subgrids:
                if include_grid[subgrid.attrib['Name']] == 0:
                    grid.remove(subgrid)
    else:
        for grid in volume:
            if include_grid[grid.attrib['Name']] == 0:
                volume.remove(grid)

    for grid in volume_and_boundaries:
        subgrids = grid.findall('.//Grid')
        if len(subgrids) == 0:
            domain.remove(grid)

    # Write out the new xmf file.
    tree.write(outputfile,
               encoding='utf-8',
               xml_declaration=True,
               pretty_print=True)



