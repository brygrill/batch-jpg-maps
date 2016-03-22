# batch-jpg-maps
Use arcpy to make maps from Excel list

- Converts xlsx files to xls so xlrd module will work
- Reads xls file and adds field to list
- Loops through list and makes map based on list
    + In this case it's a list of unique parcel IDs and a map zoomed to each parcel
- Extracts each mxd as jpg and saves to folder 