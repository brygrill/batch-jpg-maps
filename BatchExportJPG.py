#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      grillb
#
# Created:     10/06/2015
# Copyright:   (c) grillb 2015
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import glob, os, time, shutil, win32com.client, arcpy
import arcpy.mapping as MAP

#Global Variables
xlsx_files = glob.glob('path\to\files\*.xlsx')
accountList = []
root = r"\path\to\root"
outJPG = r"path\to\output"
mxdS = r"path\to\map\on\network\map.mxd"
mxdC = r"path\to\local\copy\map.mxd"

#Create Workspace Folders and Map Doc
if not os.path.exists(root): os.makedirs(root)
if not os.path.exists(outJPG): os.makedirs(outJPG)
if not os.path.exists(mxdC): shutil.copy(mxdS, mxdC)

#Convert xlsx files to xls to xlrd module will work
def convertXLSX():
    print xlsx_files

    if len(xlsx_files) == 0:
        raise RuntimeError('No XLSX files to convert.')

    xlApp = win32com.client.Dispatch('Excel.Application')

    for file in xlsx_files:
        xlWb = xlApp.Workbooks.Open(os.path.join(os.getcwd(), file))
        xlWb.SaveAs(os.path.join(os.getcwd(), file.split('.xlsx')[0] +
    '.xls'), FileFormat=1)

    xlApp.Quit()

    time.sleep(2) # give Excel time to quit, otherwise files may be locked
    for file in xlsx_files:
        os.unlink(file)


def readExcel():
    import xlrd
    xcelList = []
    workbook = xlrd.open_workbook(r"C:\TEMP\testing_xlrd.xls")
    worksheet = workbook.sheet_by_name('Sheet1')
    for curr_row in range(worksheet.nrows):
         row = worksheet.row(curr_row)
         xcelList.append(row[0:1])

    print xcelList, "\n"

    import itertools

    xlRecords = itertools.chain(*xcelList)
    xlRecords = list(xlRecords)

    for item in xlRecords:
        item = str(item)
        item = item.replace("'", "")
        accountList.append(item[6:])

    print accountList, "\n"

def makeMap():
    def noAcctLog():
        txtfilename = r"%s\NoAccount.txt" % (outJPG)
        with open(txtfilename, "a") as mytxtfile:
            mytxtfile.write("{}\n".format(acct))

    def zoomExtractMXD(parcel):
        #Variables
        mxd = MAP.MapDocument(mxdC)
        df = MAP.ListDataFrames(mxd)[0]
        parcelLyr = MAP.ListLayers(mxd, "Parcels", df)[0]
        jpg = os.path.join(outJPG, "%s") % (parcel)
        #Update Layers
        whereClause = "ACCOUNT ='%s'" % parcel
        arcpy.SelectLayerByAttribute_management(parcelLyr, "NEW_SELECTION", whereClause)
        #Make sure Account Exists
        getCount = arcpy.GetCount_management(parcelLyr)
        theCount = int(getCount.getOutput(0))
        if theCount >= 1:
            df.extent = parcelLyr.getSelectedExtent(True)
            df.scale *= 1.1
            #Update Text
            updatetext1 = MAP.ListLayoutElements(mxd, "TEXT_ELEMENT", "ACCOUNT")[0]
            updatetext1.text = parcel
            #Export to JPG
            MAP.ExportToJPEG(mxd, jpg)
            #Release MXD
            del mxd
        else:
            noAcctLog()

    for acct in accountList:
        print acct
        zoomExtractMXD(acct)

def main():
    convertXLSX()
    readExcel()
    makeMap()

if __name__ == '__main__':
    main()




