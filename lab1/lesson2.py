import os
import magic
import hashlib

import csv

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument


from PIL import Image
from PIL.ExifTags import TAGS
import docx

def getMetaData(doc):
    #https://stackoverflow.com/questions/61242017/how-to-extract-metadata-from-docx-file-using-python
    metadata = {}
    prop = doc.core_properties
    metadata["author"] = prop.author
    metadata["category"] = prop.category
    metadata["comments"] = prop.comments
    metadata["content_status"] = prop.content_status
    metadata["created"] = prop.created
    metadata["identifier"] = prop.identifier
    metadata["keywords"] = prop.keywords
    metadata["last_modified_by"] = prop.last_modified_by
    metadata["language"] = prop.language
    metadata["modified"] = prop.modified
    metadata["subject"] = prop.subject
    metadata["title"] = prop.title
    metadata["version"] = prop.version
    return metadata






path = input("Enter a path: ")

with open("output.csv", "w", newline='') as writeFile:
    writer = csv.writer(writeFile)

    for root, dir, files in os.walk(path, topdown=False):
        for file in files:
            openedFile = open(os.path.join(root,file),"rb").read()
            sha256hash = hashlib.sha256(openedFile)
            md5hash = hashlib.md5(openedFile)
            # print(md5Hash.hexdigest())
            fileInfo = magic.from_file(os.path.join(root,file))
            print(fileInfo)


            info = "Not a PDF bud!"

            if ("PDF" in fileInfo):
                fp = open(os.path.join(root,file),"rb")
                parser = PDFParser(fp)
                doc = PDFDocument(parser)
                info = doc.info


            elif ("Microsoft Word" in fileInfo):
                doc = docx.Document(os.path.join(root,file))               
                metadata_dict = getMetaData(doc)
                info = metadata_dict
            
            elif ("JPEG" in fileInfo):
                print("IM A JPEG")
                # path to the image or video
                imagename = file
                # read the image data using PIL
                image = Image.open(imagename)

                exifdata = image.getexif()

                jpegMeta = []

                for tag_id in exifdata:
                    # get the tag name, instead of human unreadable tag id
                    tag = TAGS.get(tag_id, tag_id)
                    data = exifdata.get(tag_id)
                    # decode bytes 
                    if isinstance(data, bytes):
                        data = data.decode()
                    jpegMeta.append((f"{tag:25}: {data}"))

                info = jpegMeta
                
            

            row = [str(file), str(fileInfo), str(sha256hash.hexdigest()), str(md5hash.hexdigest()), info]



            writer.writerow(row)



