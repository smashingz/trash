import os
import pynmea2 as nmea
import xlsxwriter as xls
from xlsxwriter.utility import xl_rowcol_to_cell

def find(location):
    dir_list=[]
    for root, dirs, files in os.walk(location):
        for file in files:
            if file[-3:] == "ubx":
               dir_list +=[(root, file)]
    return dir_list

def parse(file_name):
    data = []
    with open(file_name) as file:
        for line in file:
            msg = ''
            try:
                msg = nmea.parse(line)
            except nmea.ParseError:
                print("String \"%s\" is not parsed in file %s" % (line, file_name))
            if type(msg) == nmea.types.talker.GGA:
                data += [(msg.latitude, msg.longitude, msg.timestamp)]
    return data

def create_report(data, file_name, coord):
    report = xls.Workbook(file_name)
    data_sheet = report.add_worksheet("Data")
    time_format = report.add_format()
    time_format.set_num_format("hh:mm:ss")
    data_sheet.write(0, 7, "Latitude(exemplary)")
    data_sheet.write(0, 8, "Longitude(exemplary)")
    data_sheet.write(1, 7, coord[0])
    data_sheet.write(1, 8, coord[1])
    data_sheet.write(0, 0, "Latitude")
    data_sheet.write(0, 1, "Longitude")
    data_sheet.write(0, 2, "Time (present)")
    data_sheet.write(0, 3, "Time (measurement)")
    data_sheet.write(0, 4, "Inaccuracy, m")
    row, col = 1, 0
    for lat, long, time in data:
        data_sheet.write(row, col, lat)
        data_sheet.write(row, col + 1, long)
        data_sheet.write(row, col + 2, time, time_format)
        data_sheet.write(row, col + 3, "=%s-%s" % (xl_rowcol_to_cell(row, col + 2), xl_rowcol_to_cell(0, col + 2, row_abs = True, col_abs = True)), time_format)
        data_sheet.write(row, col + 4, "=IF(OR(%s=0,%s=0),0,"
                    "6371*(ACOS(SIN(RADIANS(%s))*SIN(RADIANS(%s))+"
                    "COS(RADIANS(%s))*COS(RADIANS(%s))*COS(RADIANS(%s)-RADIANS(%s)))) * 1000)" % (
                        xl_rowcol_to_cell(row, col),
                        xl_rowcol_to_cell(row, col + 1),
                        xl_rowcol_to_cell(row, col),
                        xl_rowcol_to_cell(1,7, row_abs = True, col_abs = True),
                        xl_rowcol_to_cell(row, col),
                        xl_rowcol_to_cell(1,7, row_abs = True, col_abs = True),
                        xl_rowcol_to_cell(row, col + 1),
                        xl_rowcol_to_cell(1,8, row_abs = True, col_abs = True))
                    )
        row += 1
    chart_sheet = report.add_chartsheet("Chart")
    chart = report.add_chart({"type": "scatter", "subtype": "straight"})
    chart.add_series({
        "values": ["Data", 0, 4, row - 1, 4],
        "categories": ["Data", 0, 3, row - 1, 3],
        })
    chart.set_x_axis({
        "name": "time",
         "major_gridlines": {'visible': True},
         "minor_gridlines": {'visible': True},
        })
    chart.set_y_axis({
        "name": "inaccuracy",
        "major_gridlines": {'visible': True},
        "minor_gridlines": {'visible': True},
        })
    chart_sheet.set_chart(chart)
    report.close()

files = find("./")
rez_dir = "rez"
work_dir = os.getcwd()
for file in files:
    data = parse(os.path.join(file[0], file[1]))
    rez_path = os.path.join(rez_dir, file[0])
    if not os.access(rez_path, os.R_OK):
        os.makedirs(rez_path)
    rez_file = file[1].replace("ubx", "xlsx")
    os.chdir(rez_path)
    create_report(data, rez_file, [44.591052, 33.482604])
    print("Report %s is created..." % (os.path.join(file[0], file[1])))
    os.chdir(work_dir)
print("Done!:)")
