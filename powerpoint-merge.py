#!/usr/bin/env python3.7

# Graduation presentation builder
# Paul Baumgarten 2019
import os
import sys
import json
# External packages
import pandas as pd
from openpyxl import load_workbook, Workbook
from gooey import Gooey, GooeyParser
# Project imports
from app import PowerPointer

# ********** NO CHANGES REQUIRED BELOW THIS POINT UNLESS BREAKING API CHANGE **********

def get_excel_data(xlsx_filename, worksheet):
    # Read excel file into json data
    xl = pd.read_excel(xlsx_filename, sheet_name=worksheet, index_col=None, na_values=["NA"])
    data = []
    for rowid, row in sorted(xl.iterrows()):
        record = {}
        for k,v in row.items():
            record[k] = v
        data.append(record)
    return(data)

# Begin code
@Gooey()
def main():
    # Setup via Gooey
    parser = GooeyParser(description="Powerpoint merger")

    parser.add_argument('PPT_template', metavar='Powerpoint master', help="Powerpoint template master", widget='FileChooser') 
    parser.add_argument('Slides_to_use', metavar='Layout slide IDs', help="Comma separated list of layout slide ids to parse for each record") 
    parser.add_argument('Media_folder', metavar='Media folder', help="Folder containing any images required", widget='DirChooser') 
    parser.add_argument('Excel_source', metavar='Excel file', help="Excel file containing data for merging", widget='FileChooser') 
    parser.add_argument('Excel_worksheet', metavar='Excel worksheet', help="Excel worksheet containing data for merging") 
    parser.add_argument('PPT_Save_as', metavar='Save as', help="Save Powerpoint render as", widget='FileSaver') 

    args = parser.parse_args()

    slides_per_record = []
    if args.Slides_to_use.count(", "):
        print("layout slide ids should be comma separated without spaces.")
        return

    ppt_template = args.PPT_template
    media_folder = args.Media_folder
    data_file = args.Excel_source
    worksheet = args.Excel_worksheet
    slides_per_record = args.Slides_to_use.split(",")
    ppt_output = args.PPT_Save_as

    #media_folder =  "/Users/pbaumgarten/Desktop/2020 Graduation ceremony/pictures"
    #data_file =     "/Users/pbaumgarten/Desktop/2020 Graduation ceremony/Graduation 2020 - Presentation details (Responses) (1).xlsx"
    #worksheet =     "Form responses 1"
    #ppt_template =  "/Users/pbaumgarten/Desktop/2020 Graduation ceremony/LIVE TEMPLATE.pptx"
    #ppt_output =    "/Users/pbaumgarten/Desktop/2020 Graduation ceremony/Render.pptx"
    # -> Supply the name of the "master template layout slide" you wish to parse for each record


    # Be advised:
    # * Column names in the Excel spreadsheet that you want to use within a Powerpoint merge CAN NOT have spaces/punctuation (except an underscore)




    # Get to work....
    ppt = PowerPointer.PowerPointer(ppt_template, media_folder)

    # Load student information
    recordset = get_excel_data(data_file, worksheet)

    # Get list of student image files
    media = os.listdir(media_folder)
    for i in range(len(media)):
        media[i] = os.path.join(media_folder, media[i])

    # For every row in the spreadsheet
    row_num = 0
    for record in recordset:
      #if student['studentid'] == "ST18193": # *** Use to test with only Tia Bagha  ***
        # Print progress updates so if we get errors we know which record is causing the problem
        print(f"Processing record {row_num}...")
        row_num = row_num + 1

        # Create the content for this student
        for layout_name in slides_per_record:

            # Create a new slide
            new_slide = ppt.new_slide(layout_name)

            # Fill placeholders with the content for this slide
            ppt.parse_placeholders(new_slide.slide_id, record)

    # Save resulting presentation
    print(f"Saving output PPT to: {ppt_output}")
    ppt.save(ppt_output)

if __name__ == "__main__":
    main()
