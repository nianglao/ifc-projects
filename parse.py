# python2.7

import sys

fname=sys.argv[1]
print 'Input file name ', fname

import ifcopenshell
fobj=ifcopenshell.open(fname)
project=fobj.by_type('IfcProject')[0]
print 'project ', project

