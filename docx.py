import zipfile, re
path_to_zip_file = "test.docx"
directory_to_extract_to = path_to_zip_file + ".unzip"
docx_xml = directory_to_extract_to + "/word/document.xml"

with zipfile.ZipFile(path_to_zip_file, 'r') as zip_ref:
    zip_ref.extractall(directory_to_extract_to)

# Read in the file
with open(docx_xml, 'rb') as file :
  filedata = file.read()

filedata = filedata.decode('UTF-8')  
filedata = filedata.replace('&lt;p&gt;', '')
filedata = filedata.replace('&lt;/p&gt;', '')


new_filedata = ""
regex = r"<w:t xml:space=\"preserve\">Finding XX\:"
total_finding = len(re.findall(regex, filedata, re.MULTILINE))
"""
from xml.etree import ElementTree as ET
ET.register_namespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
tree = ET.fromstring(filedata)
tree.find('body/p/hyperlink/r/t').text = '1/1/2011'
tree.write(filedata)
"""

print("Total finding : " + str(total_finding))
for finding_index in range(1, total_finding+1):
    replace_str = "<w:t xml:space=\"preserve\">Finding " + str(finding_index) + ":"
    filedata = re.sub(regex, replace_str, filedata, finding_index)
    print(finding_index)

#print(filedata)

################

#if result:
#    print(result)

# Write the file out again
with open(docx_xml, 'w', encoding="utf8") as file:
  file.write(filedata)


"""
xml_finding = re.findall(r"<w:t>Finding XX:", filedata, re.MULTILINE)
total_finding = len(xml_finding)
print(type(xml_finding))
print(xml_finding)

for num in range(1, total_finding):
    print("{:02d}".format(num)) #py3
    print("%02d" % (number,)) #py2
"""
#===============================

