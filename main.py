from docx import Document
inputFile = Document("H:\LOCK\Documents\Books\The Fate of Ten\processing.docx")
outputFile = Document()

# reading input
paragraphs = []
for p in inputFile.paragraphs:
    paragraphs.append(p.text)

# looping through paragraphs list to remove multiple line breaks
index = 0
while index < paragraphs.__len__():
    if (paragraphs[index] == " "):
        paragraphs[index - 1] += " " + paragraphs[index + 1]
        paragraphs.pop(index + 1)
        paragraphs.pop(index)
        index -= 1
    index += 1

# looping through paragraphs to remove redundant
index = 0
while index < paragraphs.__len__():
    if "\t" in paragraphs[index]:
        paragraphs[index] = paragraphs[index].replace("\t", " ")
        index -= 1
    index += 1

#print out
for p in paragraphs:
    outputFile.add_paragraph(p)
outputFile.save("H:\LOCK\Documents\Books\The Fate of Ten\outputFile.docx")






