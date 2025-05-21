import uno
import xml.etree.ElementTree as ET

def import_drawio_diagram(*_):
    # Step 1: Setup LibreOffice Draw document
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()

    if not doc.supportsService("com.sun.star.drawing.DrawingDocument"):
        raise Exception("Please run this macro inside LibreOffice Draw.")

    page = doc.getDrawPages().getByIndex(0)

    # Step 2: Load Draw.io XML file
    xml_path = "/home/molefe/.config/libreoffice/4/user/Scripts/python/shapes.drawio"  # Adjust as needed
    tree = ET.parse(xml_path)
    root = tree.getroot()

    for cell in root.findall(".//mxCell"):
        vertex = cell.attrib.get("vertex")
        edge = cell.attrib.get("edge")
        value = cell.attrib.get("value", "")
        style = cell.attrib.get("style", "")
        geometry = cell.find("mxGeometry")

        if vertex == "1" and geometry is not None:
            # Extract shape properties
            x = float(geometry.attrib.get("x", 0))
            y = float(geometry.attrib.get("y", 0))
            width = float(geometry.attrib.get("width", 10))
            height = float(geometry.attrib.get("height", 5))

            # Determine shape type from style
            if "ellipse" in style:
                shape = doc.createInstance("com.sun.star.drawing.EllipseShape")
            else:
                shape = doc.createInstance("com.sun.star.drawing.RectangleShape")

            # Set size and position
            shape.setSize(uno.createUnoStruct("com.sun.star.awt.Size", int(width * 10), int(height * 10)))
            shape.setPosition(uno.createUnoStruct("com.sun.star.awt.Point", int(x * 10), int(y * 10 )))
            shape.setString(value)
            page.add(shape)

        elif edge == "1":
            # (Optional) Handle connectors
            pass

