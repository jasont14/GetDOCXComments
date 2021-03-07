#Jason Thatcher
#2021-03-07

#reference https://python-docx.readthedocs.io/en/latest/api/document.html#id1
#Module https://lxml.de/3.7/index.html
#namespace reference: http://www.datypic.com/sc/ooxml/ns-w.html
#reference https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing?view=openxml-2.8.1
#reference https://stackoverflow.com/  -- Several DOCX Q&A

from docx import Document
from lxml import etree
import zipfile, os

ooXMLns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
global dir_root

#get filenames from directory
#returns list of files with docx extension
def get_docx_files():
    global dir_root
    dir_root = input("Enter Directory: ")
    return (dir_root + "\\" + x for x in os.listdir(dir_root) if x.endswith('.docx'))
    
#open each docx file to get comments and write to csv
def loop_through_docx():
    for f in get_docx_files():
        comm_string = comments_with_reference_paragraph(f)
        write_lines_to_file(comm_string)

#write comments to CSV.  Include FileName, Comment, CommentReferenceText
def write_lines_to_file(lines):
    global dir_root
    out_dir = dir_root + "\\filecomments.txt"
    with open(out_dir, 'a+') as fo:
        for line in lines:
            fo.write('%s\n' % line)

#Uses LXML module and ZIPFILE. Returns dictionary id,comment
def get_document_comments(docxFileName):
    comments_dict={}
    docxZip = zipfile.ZipFile(docxFileName)
    commentsXML = docxZip.read('word/comments.xml')
    et = etree.XML(commentsXML)
    comments = et.xpath('//w:comment',namespaces=ooXMLns)
    for c in comments:
        comment=c.xpath('string(.)',namespaces=ooXMLns)
        comment_id=c.xpath('@w:id',namespaces=ooXMLns)[0]
        comments_dict[comment_id]=comment
    return comments_dict

#Uses python-docx module
#Function to fetch all the comments in a paragraph
def paragraph_comments(paragraph,comments_dict):
    comments=[]
    for run in paragraph.runs:
        comment_reference=run._r.xpath("./w:commentReference")
        if comment_reference:
            comment_id=comment_reference[0].xpath('@w:id',namespaces=ooXMLns)[0]
            comment=comments_dict[comment_id]
            comments.append(comment)
    return comments

#Get the paragraphs in document and match to comments if any.
#returns list of comments as CSV
def comments_with_reference_paragraph(docxFileName):

    document = Document(docxFileName)
    comments_dict = get_document_comments(docxFileName)

    document_title = os.path.basename(docxFileName)
    document_comment = ""
    document_comment_reference = ""
    comment_string = []
    
    for paragraph in document.paragraphs: 

        if comments_dict: 
            comments=paragraph_comments(paragraph,comments_dict)  
        
            if comments:
                document_comment = "".join(comments[0])
                document_comment_reference = paragraph.text
                #comments_with_their_reference_paragraph.append({paragraph.text: comments})
    
                comment_string.append(document_title + "," + document_comment + ",\"" +  document_comment_reference + "\"")
                # print("Title: ", document_title)
                # print("Comment: ", document_comment)
                # print("Comment Reference: ", document_comment_reference)
    return comment_string

if __name__=="__main__":    
    loop_through_docx()