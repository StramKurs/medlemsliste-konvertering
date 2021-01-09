import sys

sys.path.append('../src/medlemsliste')

import pytest
import memberlist

class Test():

    

    def test_cvs_import(self):
        app = memberlist.convertCvs()
        app.run()