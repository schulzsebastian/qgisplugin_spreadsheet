import os
from pyexcel_ods import ods
from base import ODSCellTypes


class TestODSReader(ODSCellTypes):
    def setUp(self):
        r = ods.ODSBook()
        r.open(os.path.join("tests",
                            "fixtures",
                            "ods_formats.ods"))
        self.data = r.read_all()
        for key in self.data.keys():
            self.data[key] = list(self.data[key])
        r.close()


class TestODSWriter(ODSCellTypes):
    def setUp(self):
        r = ods.ODSBook()
        r.open(os.path.join("tests",
                            "fixtures",
                            "ods_formats.ods"))
        self.data1 = r.read_all()
        self.testfile = "odswriter.ods"
        w = ods.ODSWriter()
        w.open(self.testfile)
        w.write(self.data1)
        w.close()
        r2 = ods.ODSBook()
        r2.open(self.testfile)
        self.data = r2.read_all()
        for key in self.data.keys():
            self.data[key] = list(self.data[key])

    def tearDown(self):
        if os.path.exists(self.testfile):
            os.unlink(self.testfile)
