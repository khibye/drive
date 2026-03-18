from spire.doc import Document

doc = Document()
doc.LoadFromFile("doc.docx")

for sec_idx in range(doc.Sections.Count):
    section = doc.Sections.get_Item(sec_idx)
    for para_idx in range(section.Paragraphs.Count):
        para = section.Paragraphs.get_Item(para_idx)
        print(repr(para.Text))

doc.Close()
