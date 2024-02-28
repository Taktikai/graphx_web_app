import xml.etree.ElementTree as ET
import unicodedata
from collections import deque

#OBJEKTUMOK########################################################################################################################
###################################################################################################################################
class Device:
    def __init__(self, id, description, típus, elnevezés, SN, saját_cím, fogadó_eszköz, kommunikáció, fogadó_interface, hálózatazonosító_frekvencia):
        self.id = id
        self.description = description
        self.típus = típus
        self.elnevezés = elnevezés
        self.SN = SN
        self.saját_cím = saját_cím
        self.fogadó_eszköz = fogadó_eszköz
        self.kommunikáció = kommunikáció
        self.fogadó_interface = fogadó_interface
        self.hálózatazonosító_frekvencia = hálózatazonosító_frekvencia
        self.connected_devices = {}

class Group:
    def __init__(self, id, típus, elnevezés, closed = False):
        self.id = id
        self.típus = típus
        self.elnevezés = elnevezés
        self.closed = closed
        self.content = {}


#FÜGGVÉNYEK###########################################################################################################################
######################################################################################################################################
def ignorator_comparator(str1, str2):
    str1_normalized = unicodedata.normalize('NFKD', str1).encode('ASCII', 'ignore').decode('utf-8')
    str2_normalized = unicodedata.normalize('NFKD', str2).encode('ASCII', 'ignore').decode('utf-8')

    str1_normalized = str1_normalized.replace(" ", "")
    str2_normalized = str2_normalized.replace(" ", "")

    str1_normalized = str1_normalized.replace("_", "")
    str2_normalized = str2_normalized.replace("_", "")

    str1_normalized = str1_normalized.replace("-", "")
    str2_normalized = str2_normalized.replace("-", "")

    return str1_normalized.lower() == str2_normalized.lower()


#PARSING########################################################################################################################
################################################################################################################################
def get_node_keys(keys):
    key_ids = []
    for key in keys:
        if key.get('attr.name') and not ignorator_comparator(key.get('attr.name'), 'url') and not None and key.get("for") == "node":
            items = [key.get('id'),key.get('attr.name')]
            key_ids.append(items)
    return key_ids

def get_edge_keys(keys):
    key_ids = []
    for key in keys:
        if key.get('attr.name') and not ignorator_comparator(key.get('attr.name'), 'url') and not None and key.get("for") == "edge":
            items = [key.get('id'),key.get('attr.name')]
            key_ids.append(items)
    return key_ids


def get_element_depth(root, element_id, current_depth=0):
    if root.get('id') == element_id:
        return current_depth
    for child in root:
        depth = get_element_depth(child, element_id, current_depth + 1)
        if depth is not None:
            return depth
    return None


def get_specific_data(node, data_name, keys):
    for key in keys:
        if ignorator_comparator(key[1], data_name):
            my_id = key[0]        
    for data in node.iter("{http://graphml.graphdrawing.org/xmlns}data"):
        if data.get('key') == my_id:
            return data.text
        elif data.get('key') == "d10" or data.get('key') == "d17":
            return None


def reverse_search(item, filename, root_group, id = None):
    if isinstance(item, Device):
        root = ET.parse(filename).getroot()
        for node in root.iter("{http://graphml.graphdrawing.org/xmlns}node"):
            if item.id == node.get('id'):
                return node
    elif id == None:
        id = item.get('id')
        queue = deque([root_group])
        while queue:
            group = queue.popleft()
            if id in group.content and isinstance(group.content[id], Device):
                return group.content[id]
            for item in group.content.values():
                if isinstance(item, Group):
                    queue.append(item)
    elif item == None:
        queue = deque([root_group])
        while queue:
            group = queue.popleft()
            if id in group.content and isinstance(group.content[id], Device):
                return group.content[id]
            for item in group.content.values():
                if isinstance(item, Group):
                    queue.append(item)
    return None
        


#STRUKTÚRA KIALAKÍTÁS################################################################################################################
######################################################################################################################################
def comms(edge, keys):
    if edge.find(".//y:GenericEdge", namespaces={"y": "http://www.yworks.com/xml/graphml"}) is not None:
                    if edge.find(".//y:LineStyle", namespaces={"y": "http://www.yworks.com/xml/graphml"}).get("type") == "dashed":
                                return "wifi" 
                    if edge.find(".//y:LineStyle", namespaces={"y": "http://www.yworks.com/xml/graphml"}).get("type") == "dashed_dotted":
                                return "gsm"
                    if edge.find(".//y:LineStyle", namespaces={"y": "http://www.yworks.com/xml/graphml"}).get("type") == "line":
                                return "ethernet"
    if edge.find(".//y:LineStyle", namespaces={"y": "http://www.yworks.com/xml/graphml"}).get("type") == "dashed":
        return "rádiós soros port" 
    if edge.find(".//y:LineStyle", namespaces={"y": "http://www.yworks.com/xml/graphml"}).get("type") == "dashed_dotted":
        return "rádiós impulzus"
    if edge.find(".//y:LineStyle", namespaces={"y": "http://www.yworks.com/xml/graphml"}).get("type") == "dotted":
        return "LoRa"
    if get_specific_data(edge, "Kommunikáció Módja", get_edge_keys(keys)) is not None:
        return get_specific_data(edge, "Kommunikáció Módja", get_edge_keys(keys))        


def get_all_devices(filename):
    tree = ET.parse(filename)
    root = tree.getroot()
    keys = root.findall(".//{http://graphml.graphdrawing.org/xmlns}key")
    root_group = Group('root', 'root', filename[:-8])
    groups = {'root': root_group}
    for node in root.iter('{http://graphml.graphdrawing.org/xmlns}node'):
        if 'yfiles.foldertype' in node.attrib and (node.attrib['yfiles.foldertype'] == 'group' or node.attrib['yfiles.foldertype'] == 'folder'):
            group = Group(id=node.get("id"),
                        típus=get_specific_data(node, "Típus", get_node_keys(keys)),
                        elnevezés=get_specific_data(node, "Elnevezés", get_node_keys(keys)),)
            if node.attrib['yfiles.foldertype'] == 'folder':
                group.closed = True
            groups[node.get('id')] = group
            if len(node.get('id')) > 5:
                parent_id = node.get('id').rsplit('::', 1)[0]
                if parent_id in groups:
                    groups[parent_id].content[group.id] = group
            else:
                parent_id = "root"
                if parent_id in groups:
                    groups[parent_id].content[group.id] = group
        elif get_specific_data(node, "description", get_node_keys(keys)) is not None:
            device = Device(id=node.get("id"), 
                            description=get_specific_data(node, "description", get_node_keys(keys)),
                            típus=get_specific_data(node, "Típus", get_node_keys(keys)),
                            elnevezés=get_specific_data(node, "Elnevezés", get_node_keys(keys)),
                            SN=get_specific_data(node, "SN", get_node_keys(keys)),
                            saját_cím=get_specific_data(node, "Saját Cím", get_node_keys(keys)),
                            fogadó_eszköz=None,
                            kommunikáció=None,
                            fogadó_interface=None,
                            hálózatazonosító_frekvencia=None,)
            if len(node.get('id')) > 2:
                parent_id = node.get('id').rsplit('::', 1)[0]
                if parent_id in groups:
                    groups[parent_id].content[device.id] = device
            else:
                parent_id = "root"
                if parent_id in groups:
                    groups[parent_id].content[device.id] = device
    connections(filename, root_group)
    close_groups(root_group)
    return root_group


def connections(filename, root_group):
    root = ET.parse(filename).getroot()
    edges = root.findall(".//{http://graphml.graphdrawing.org/xmlns}edge")
    keys = root.findall(".//{http://graphml.graphdrawing.org/xmlns}key")
    for edge in edges:
        parent = reverse_search(None, filename, root_group, edge.get("target"))
        device = reverse_search(None, filename, root_group, edge.get("source"))
        if device is not None and parent is not None:
            device.fogadó_eszköz=get_specific_data(reverse_search(parent, filename, root_group), "Típus", get_node_keys(keys))
            device.kommunikáció=comms(edge, keys)
            device.fogadó_interface=get_specific_data(edge, "Fogadó interface", get_edge_keys(keys))
            device.hálózatazonosító_frekvencia=get_specific_data(edge, "Hálózatazonosító/Frekvencia", get_edge_keys(keys))
            if parent.típus and "remx" in parent.típus.lower() and device.description == "Áramváltó":
                device.típus += "/50A"
            parent.connected_devices[device.id] = device


def close_groups(root_group, closed = False):
    for item in root_group.content.values():
        if isinstance(item, Group):
            if closed or item.closed:
                item.closed = True
                close_groups(item, closed=True)
            else:
                close_groups(item)
    return


def count_devices_by_type(group, filename):
    counts = {}
    for id, item in group.content.items():
        if isinstance(item, Device):
            if item.típus == None:
                item.típus = item.description
            if item.típus not in counts:
                counts[item.típus] = 0
            if item.description == "Áramváltó":
                node = reverse_search(item, filename, group)
                fill_color = node.find(".//y:Fill", namespaces={"y": "http://www.yworks.com/xml/graphml"}).get("color")
                if fill_color is not None and len(fill_color) > 7 and fill_color[:8] == '#0000000':
                        counts[item.típus] += int(fill_color[-1])
                else:
                    counts[item.típus] += 1
            else:    
                counts[item.típus] += 1
        elif isinstance(item, Group):
            sub_counts = count_devices_by_type(item, filename)
            for típus, count in sub_counts.items():
                if típus not in counts:
                    counts[típus] = 0
                counts[típus] += count
            if not any(isinstance(sub_item, Group) for sub_item in item.content.values()):
                if item.típus not in counts:
                    counts[item.típus] = 0
                counts[item.típus] += 1
    return counts