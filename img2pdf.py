from PIL import Image
import os


imagePath = r'C:\Users\921722\Downloads\(New submission same as TCAJV with TGC format)_MS for Loose Soil Permeation Grouting_REV.04-'
imagelist = os.listdir(imagePath)

for image in imagelist:
    image_1 = Image.open(f'C:/Users/921722/Downloads/(New submission same as TCAJV with TGC format)_MS for Loose Soil Permeation Grouting_REV.04-/{image}')
    im_1 = image_1.convert('RGB')
    image_name = image.replace(".pdf", "")
    im_1.save(f'img_AD\{image_name}.pdf')


