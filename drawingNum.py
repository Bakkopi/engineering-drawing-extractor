import cv2
import pytesseract
from matplotlib import pyplot as pt
import numpy as np

pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
# Page segmentation mode, PSM was changed to 6 since each page is a single uniform text block.
    
def GetString(img, keyword1, keyword2):
    copy = img.copy()
    [nrow, ncol]= img.shape
    
    blur = cv2.GaussianBlur(copy, (3,3), 0)
    ret, thresh = cv2.threshold(blur, 127, 1, cv2.THRESH_BINARY_INV)
    
    contours, hierarchy = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    
    for contour in contours:
        area = cv2.contourArea(contour)
        
        if (area > 40000 and area < 5000000):
            x,y,w,h = cv2.boundingRect(contour)
            cv2.rectangle(img, (x,y), (x+w, y+h), (36,255,12), -1)
            ROI = copy[y:y+h, x:x+w]
            
            string = (pytesseract.image_to_string(ROI, config ='--psm 6')).strip()
            if (string == ""):
                return
            
            # pt.figure()
            # pt.imshow(ROI, cmap = "gray")
            
            if (keyword1 in string or keyword2 in string):
                ROI = copy[y:y+h+100, x+10:x+w]  # we take a larger area of the box identified   
    
                copyROI = ROI.copy()
                ret, thresh = cv2.threshold(ROI, 0, 255, cv2.THRESH_BINARY_INV)
                
                #--- Remove any potential boxes surrounding the letters which can impair extraction through OCR ---#
                # Remove horizontal lines
                horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40,1))
                remove_horizontal = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)
                cnts = cv2.findContours(remove_horizontal, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                cnts = cnts[0] if len(cnts) == 2 else cnts[1]
                for c in cnts:
                    cv2.drawContours(copyROI, [c], -1, (255,255,255), 5)
                
                # Remove vertical lines
                vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1,30))
                remove_vertical = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, vertical_kernel, iterations=2)
                cnts = cv2.findContours(remove_vertical, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                cnts = cnts[0] if len(cnts) == 2 else cnts[1]
                for c in cnts:
                    cv2.drawContours(copyROI, [c], -1, (255,255,255), 5)
                
                # --- Final Reading of Box --- #
                string = (pytesseract.image_to_string(copyROI, config ='--psm 6')).strip()    
                string = string.splitlines()
                extracted_string = ""

                for i in range(len(string)):
                    
                    if keyword1 in string[i] or keyword2 in string[i]:
                        indexOfValue = i
                        while extracted_string == "":     
                            indexOfValue = indexOfValue + 1
                            if ((indexOfValue) < len(string)): # if true means that this string has this index
                                extracted_string = string[indexOfValue]
                            else:
                                break
                        return extracted_string
                
                return extracted_string
                break
   
# img = cv2.imread("08.png", 0)     
# data_extract = {}     
  
# drawingNum = GetString(img, "DRAWING NUMBER", "DRAWING NO")
# drawnBy = GetString(img, "DRAWN BY", "DRAWN")
# checkedBy = GetString(img, "CHECKED BY", "CHECKED")
# title = GetString(img, "TITLE", "DRAWING TITLE")
# approvedBy = GetString(img, "APPPROVED BY", "APPROVED")
# contractor = GetString(img, "CONTRACTOR", "COMPANY")
# unit = GetString(img, "UNIT", "UNIT")
# status = GetString(img, "STATUS", "STATUS")
# page = GetString(img, "PAGE", "PAGE")
# projectNum = GetString(img, "PROJECT NO", "PROJECT NUM")
# lang = GetString(img, "LANG", "LANG")
# cad = GetString(img, "CAD NO", "CAD")
# font = GetString(img, "FONT", "FONT STYLE")

# data_extract["drawing number"] = drawingNum
# data_extract["drawn by"] = drawnBy
# data_extract["checked by"] = checkedBy
# data_extract["title"] = title
# data_extract["approved by"] = approvedBy
# data_extract["contractor"] = contractor
# data_extract["unit"] = unit
# data_extract["status"] = status
# data_extract["page"] = page
# data_extract["projectNum"] = projectNum
# data_extract["lang"] = lang
# data_extract["cad"] = cad
# data_extract["font"] = font

# print(data_extract)