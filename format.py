from lxml import etree

doc = etree.parse('./shed.xml')
s = etree.tostring(doc, pretty_print=True, encoding=str)
o = open('./shed1.xml', "w")
o.write(s)
o.close()
