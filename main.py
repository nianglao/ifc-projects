# encoding=utf8
# python2.7

from project_tree import ProjectTree
from utils import init_logger
import csv
import codecs
from ms_project import MSProject

import sys
#reload(sys)
#sys.setdefaultencoding('utf8')

if __name__ == '__main__':
    ifc_file_path = sys.argv[1]
    ifc_csv_path = sys.argv[2]

    logger = init_logger('LOG.log')

    # get translation from ifc entity to chinese
    translation_dict = dict()
    with open('resources/ifc_translation.csv', encoding='utf-8') as f:
        for row in csv.reader(f):
            if len(row) >= 2 and len(row[1]) > 0:
                translation_dict[row[0]] = row[1]
    logger.info('translation_dict: %s', translation_dict)
    pt = ProjectTree(logger, ifc_file_path)
    # pt.print_tree(print_id=False, print_type=False,
    #              translation_dict=translation_dict)
    pt.print_microsoft_project_csv(
        translation_dict=translation_dict, out_file='yulan-project.csv')
    
    project = MSProject()
    project.convert_csv_to_project(ifc_csv_path, 'yulan.mpp')
