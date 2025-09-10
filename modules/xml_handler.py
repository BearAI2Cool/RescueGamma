import xml.etree.ElementTree as ET
from typing import Dict, List, Optional


class XMLHandler:
    def __init__(self):
        self.namespaces = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }

        for prefix, uri in self.namespaces.items():
            ET.register_namespace(prefix, uri)

    def load_xml(self, file_path: str) -> ET.ElementTree:
        return ET.parse(file_path)

    def save_xml(self, tree: ET.ElementTree, file_path: str):
        tree.write(file_path, encoding='utf-8', xml_declaration=True)

    def find_text_runs(self, tree: ET.ElementTree) -> List[ET.Element]:
        root = tree.getroot()
        return root.findall('.//a:r', self.namespaces)

    def create_gradient_fill(self, gradient_config: List[Dict]) -> ET.Element:
        grad_fill = ET.Element(f"{{{self.namespaces['a']}}}gradFill")
        grad_fill.set('flip', 'none')
        grad_fill.set('rotWithShape', '1')

        gs_list = ET.SubElement(grad_fill, f"{{{self.namespaces['a']}}}gsLst")

        for config in gradient_config:
            gs = ET.SubElement(gs_list, f"{{{self.namespaces['a']}}}gs")
            gs.set('pos', str(config['position']))

            if config['color'].startswith('#'):
                color_elem = ET.SubElement(gs, f"{{{self.namespaces['a']}}}srgbClr")
                color_elem.set('val', config['color'][1:])
            elif config['color'].startswith('accent'):
                color_elem = ET.SubElement(gs, f"{{{self.namespaces['a']}}}schemeClr")
                color_elem.set('val', config['color'])
                lum_mod = ET.SubElement(color_elem, f"{{{self.namespaces['a']}}}lumMod")
                lum_mod.set('val', '45000')
                lum_off = ET.SubElement(color_elem, f"{{{self.namespaces['a']}}}lumOff")
                lum_off.set('val', '55000')
            else:
                color_elem = ET.SubElement(gs, f"{{{self.namespaces['a']}}}srgbClr")
                color_elem.set('val', config['color'])

        lin = ET.SubElement(grad_fill, f"{{{self.namespaces['a']}}}lin")
        lin.set('ang', '0')
        lin.set('scaled', '1')

        tile_rect = ET.SubElement(grad_fill, f"{{{self.namespaces['a']}}}tileRect")

        return grad_fill

    def apply_gradient_to_text_run(self, text_run: ET.Element, gradient_config: List[Dict], font_size: str,
                                   font_name: Optional[str] = None):
        rpr = text_run.find(f"a:rPr", self.namespaces)
        if rpr is None:
            rpr = ET.SubElement(text_run, f"{{{self.namespaces['a']}}}rPr")

        rpr.set('sz', str(int(float(font_size) * 100)))

        if font_name:
            for font_type in ['latin', 'ea', 'cs']:
                font_elem = rpr.find(f"a:{font_type}", self.namespaces)
                if font_elem is None:
                    font_elem = ET.Element(f"{{{self.namespaces['a']}}}{font_type}")
                    rpr.append(font_elem)
                font_elem.set('typeface', font_name)

        for fill_elem in rpr.findall(f"a:solidFill", self.namespaces):
            rpr.remove(fill_elem)
        for fill_elem in rpr.findall(f"a:gradFill", self.namespaces):
            rpr.remove(fill_elem)

        grad_fill = self.create_gradient_fill(gradient_config)
        rpr.insert(0, grad_fill)

    def apply_gradient_to_end_para(self, paragraph: ET.Element, gradient_config: List[Dict], font_size: str = None):
        end_para_rpr = paragraph.find(f"a:endParaRPr", self.namespaces)
        if end_para_rpr is not None:
            if font_size:
                end_para_rpr.set('sz', str(int(float(font_size) * 100)))

            for fill_elem in end_para_rpr.findall(f"a:solidFill", self.namespaces):
                end_para_rpr.remove(fill_elem)
            for fill_elem in end_para_rpr.findall(f"a:gradFill", self.namespaces):
                end_para_rpr.remove(fill_elem)
            grad_fill = self.create_gradient_fill(gradient_config)
            end_para_rpr.insert(0, grad_fill)
