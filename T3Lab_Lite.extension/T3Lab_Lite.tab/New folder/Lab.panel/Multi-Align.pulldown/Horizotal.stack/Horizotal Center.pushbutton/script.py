"""


Smart Align
Smart Align is part of PyRevit Plus Optional Extensions for PyRevit

--------------------------------------------------------
pyRevit Notice:
pyRevit: repository at https://github.com/eirannejad/pyRevit

"""

__author__ = a
__version__ = b
__doc__ = 'Align Elements Horizontally: Center'

import sys
import os
sys.path.append(os.path.dirname(__file__))
from smartalign.align import main
from smartalign.core import Alignment, VERBOSE

ALIGN = Alignment.HCENTER
# ALIGN = Alignment.HLEFT
# ALIGN = Alignment.HRIGHT
# ALIGN = Alignment.VCENTER
# ALIGN = Alignment.VTOP
# ALIGN = Alignment.VBOTTOM

main(ALIGN)
