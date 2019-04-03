# encoding=utf8
# python2.7

# support IFC2X2_FINAL 

import ifcopenshell
import sys

reload(sys)
sys.setdefaultencoding('utf8')


def print_obj(level, obj):
    if level == 0:
        print obj.LongName
    else:
        print ','*level, obj.LongName
    # print '--'*level, obj.id(), obj.Name, obj.LongName, obj.is_a()


def get_children_from_aggr(obj):
    children = []
    for aggr in obj.IsDecomposedBy:
        for c in aggr.RelatedObjects:
            children.append(c)
    return children


def traverse(obj, cur_lvl=0, max_lvl=0):
    print_obj(cur_lvl, obj)
    if cur_lvl == max_lvl:
        return
    for child in get_children_from_aggr(obj):
        traverse(child, cur_lvl+1, max_lvl)


'''
project
    |-- site
        |-- building
            |-- buildingstorey
                |-- space
'''

fname = sys.argv[1]
fobj = ifcopenshell.open(fname)
projects = fobj.by_type('IfcProject')

print 'project, site, building, buildingstorey, space'
for p in projects:
    traverse(p, 0, 5)
