# encoding=utf8
# python2.7

import ifcopenshell

class TreeNode(object):
    def __init__(self, id, children = None):
        self.id = id
        self.type = ''
        self.all_ids = []
        if children is None:
            self.children = []
        else:
            self.children = children
    
    def count(self):
        return len(self.all_ids)

    def has_child(self):
        return len(self.children) > 0

    def add_child(self, child):
        self.children.append(child)
    
    def add_children(self, children):
        for c in children:
            self.children.append(c)
    
class ProjectTree(object):
    def __init__(self, logger, ifc_file_name):
        self.logger = logger
        self.file_obj = ifcopenshell.open(ifc_file_name)

        self.entities = dict()
        for id in self.file_obj.wrapped_data.entity_names():
            self.entities[id] = self.file_obj.by_id(id)
    
        projects=self.file_obj.by_type('IfcProject')
        if len(projects) != 1:
            raise Exception('Project nodes (%s) != 1'%len(projects))
        self.root=TreeNode(projects[0].id())
        self.build_project_tree(self.root)

    def get_decomposed_entity_ids(self, entity_id):
        self.logger.info('In get_decomposed_entity_ids, entity_id=%s', entity_id)
        children = []
        entity=self.entities[entity_id]
        if not hasattr(entity, 'IsDecomposedBy'):
            return
        for aggr in entity.IsDecomposedBy:
            for c in aggr.RelatedObjects:
                children.append(c.id())
        self.logger.info('children %s'%children)
        return children
    
    def get_contained_entity_ids(self, entity_id):
        self.logger.info('In get_contained_entity_ids, entity_id=%s', entity_id)
        children = []
        entity=self.entities[entity_id]
        if not hasattr(entity, 'ContainsElements'):
            return children
        for ce in entity.ContainsElements:
            for e in ce.RelatedElements:
                children.append(e.id())
        self.logger.info('children %s'%children)
        return children

    def build_project_tree(self, root):
        child_ids=[]
        if self.entities[root.id].is_a() in ['IfcProject', 'IfcSite', 'IfcBuilding']: # 'IfcBuildingStorey', 'IfcSpace'
            child_ids = self.get_decomposed_entity_ids(root.id)
            for id in child_ids:
                root.add_child(TreeNode(id))
            for c in root.children:
                self.build_project_tree(c)
        else:
            child_ids = self.get_contained_entity_ids(root.id)
            entity_dict=dict()
            # classify by entity_type
            for id in child_ids:
                entity_type = self.entities[id].is_a()
                if entity_type not in entity_dict:
                    entity_dict[entity_type] = set()
                entity_dict[entity_type].add(id)
            # build child node 
            for entity_type in entity_dict.keys():
                id_list = list(entity_dict[entity_type])
                node = TreeNode(id_list[0])
                node.all_ids = id_list
                node.type = entity_type
                root.add_child(node)

    def get_entity_name(self, id):
        entity = self.entities[id]
        name = None
        if hasattr(entity, 'LongName') and entity.LongName != None and entity.LongName != '':
            name = entity.LongName
        if name == None and hasattr(entity, 'Name') and entity.Name != None and entity.Name != '':
            name = entity.Name
        if name == None:
            name = 'Unknown'
        return name

    def print_tree(self, print_id=False, print_type=False, translation_dict=dict()):
        self._print_tree(self.root, 0, print_id, print_type, translation_dict)

    def _print_tree(self, node, level=0, print_id=False, print_type=False, translation_dict=dict()):
        attrbutes=[]
        if node.has_child():
            attrbutes.append(self.get_entity_name(node.id))
        else:
            if node.type in translation_dict:
                attrbutes.append(translation_dict[node.type])
            else:
                attrbutes.append(node.type)
            attrbutes.append(str(len(node.all_ids)))
        if print_id:
            attrbutes.append(str(node.id))
        if print_type:
            attrbutes.append(self.entities[node.id].is_a())
        print ' '*level, ','.join(attrbutes)
        for c in node.children:
            self._print_tree(c, level+1, print_id=print_id, print_type=print_type, translation_dict=translation_dict)