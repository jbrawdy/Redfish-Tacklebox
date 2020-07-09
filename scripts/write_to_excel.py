import xlsxwriter
from datetime import datetime



def write_system_inventory( inventory_list, details = False, skip_absent = False ):
    """
    Write the system inventory list into a spreadsheet

    Args:
        inventory_list: The inventory list to print
        details: True to print all of the detailed info
        skip_absent: True to skip printing absent components
    """

    # Excel workbook to save data extracted and parsed
    workbook = xlsxwriter.Workbook("./Device_Firmware_"
                                   + str('{:%Y-%m-%d_%H-%M-%S}'.format(
                                         datetime.now()))
                                   + ".xlsx")

    worksheet = workbook.add_worksheet("Device Inventory")
    cell_header_format = workbook.add_format({'bold': True, 'bg_color': 'yellow'})
    cell_name_format = workbook.add_format({'bold': True})

    column = 0
    row = 0

    # Adds header to Excel file
    header = ["NAME", "DESCRIPTION", "MANUFACTURER", "MODEL", "SKU", "PART NUMBER", "SERIAL NUMBER", "ASSET TAG" ]
    for column_title in header:
        worksheet.write(row, column, column_title, cell_header_format)
        column += 1
    row = 1

    

    
    for chassis in inventory_list:
        # Go through each component type in the chassis
        type_list = [ "Chassis", "Processors", "Memory", "Drives", "PCIeDevices", "StorageControllers", "NetworkAdapters" ]
        for inv_type in type_list:
            # Go through each component and prints its info
            for item in chassis[inv_type]:
                column = 0
                worksheet.write(row, column, inv_type, cell_name_format) 
                column += 1
                detail_list = [ "Description", "Manufacturer", "Model", "SKU", "PartNumber", "SerialNumber", "AssetTag" ]
                for detail in detail_list:
                    worksheet.write(row, column, item[detail] ) 
                    column += 1
                row += 1
    
    workbook.close()