import json
import xlwings as xw
from xlwings.utils import hex_to_rgb, rgb_to_int
import math
from sympy import Symbol,nsolve,exp
import numpy as np
import re
dataPattern = re.compile("Input data \(([A-Z]+\d+):\s*([A-Z]+\d+)\)")
_offset = (0,0)
_sheet = None

def rel_addr(x,y):
    if _sheet:
        return _sheet.range((y+_offset[1],x+_offset[0])).get_address(False,False)
    return "Address Error"

# Print iterations progress
def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
    # Print New Line on Complete
    if iteration == total: 
        print()

@xw.func
@xw.arg('data', 'range', doc='Input data range')
def subdata (data):
    return "Input data (" + data.get_address(False,False) +  ")"

@xw.func
def test (x,y):
    global _sheet
    _sheet = xw.Book.caller().sheets[0]
    return rel_addr(x,y)


#@xw.sub
def main():
    wb = xw.Book.caller()
    sheet = wb.sheets.active
    global _sheet
    _sheet = sheet
    table = sheet.used_range.value
    tl = (0, 0)
    br = (0, 0)
    found = None

    #find input data
    for line in table:
        for cell in line:
            found = re.search(dataPattern, str(cell))
            if found:
                addr_tl = str(found.group(1))
                addr_br = str(found.group(2))
                tl = (sheet[addr_tl].column, sheet[addr_tl].row)
                br = (sheet[addr_br].column, sheet[addr_br].row)
                break
        if found:
            break
    
    if tl == (0,0) or br[0]-tl[0] != 15:
        return

    num_units = br[1]-tl[1]+1

    #build address lists
    unit, stage, age_bottom, age_top, depth, thickness, mid, porosity_0, porosity_p, coeff, density_g, density_w, density_a, waterdepth, sealevel, lithology = ([] for _ in range(16))
    x = tl[0]
    for y in range(tl[1],br[1]+1):
        unit.append(sheet.range((y,x)).get_address(False,False))
        stage.append(sheet.range((y,x+1)).get_address(False,False))
        age_bottom.append(sheet.range((y,x+2)).get_address(False,False))
        age_top.append(sheet.range((y,x+3)).get_address(False,False))
        depth.append(sheet.range((y,x+4)).get_address(False,False))
        thickness.append(sheet.range((y,x+5)).get_address(False,False))
        mid.append(sheet.range((y,x+6)).get_address(False,False))
        porosity_0.append(sheet.range((y,x+7)).get_address(False,False))
        porosity_p.append(sheet.range((y,x+8)).get_address(False,False))
        coeff.append(sheet.range((y,x+9)).get_address(False,False))
        density_g.append(sheet.range((y,x+10)).get_address(False,False))
        density_w.append(sheet.range((y,x+11)).get_address(False,False))
        density_a.append(sheet.range((y,x+12)).get_address(False,False))
        waterdepth.append(sheet.range((y,x+13)).get_address(False,False))
        sealevel.append(sheet.range((y,x+14)).get_address(False,False))
        lithology.append(sheet.range((y,x+15)).get_address(False,False))

    #create total subsidence block
    printProgressBar(0, num_units, prefix = 'Total Subsidence:', suffix = 'Complete', length = 50)
    ox, oy = tl[0], br[1]+3
    linenum = 0
    global _offset
    _offset = (ox,oy)

    total_sub = []
    total_sub.append(["Total Subsidence",                 "",         "",    "km", "control",        "km",        "km",    "km"])
    sheet.range(rel_addr(0,linenum) + ":" + rel_addr(1,linenum)).merge()
    sheet.range(rel_addr(0,linenum)).font.bold = True
    sheet.range(rel_addr(0,linenum)).color = "#FFFF00"
    linenum = linenum + 1
    
    total_sub.append([                "", "Initial Porostiy", "Porosity", "mid-D",   "mid-T", "mid T x 2", "Thickness", "Depth"])
    sheet.range(rel_addr(0,linenum) + ":" + rel_addr(7,linenum)).color = "#66FFFF"
    sheet.range(rel_addr(0,linenum) + ":" + rel_addr(6,linenum)).api.Borders.Weight = 2
    sheet.range(rel_addr(7,linenum)).api.Borders.Weight = 3
    linenum = linenum + 1

    age_depth = [[""] * (num_units + 2) for _ in range(num_units+1)]
    dec_thickness = ["" for _ in range(num_units)]
    dec_depth = ["" for _ in range(num_units)]
    tec_subsidence = ["" for _ in range(num_units)]

    total_sub_line = []
    for top_unit_i in range(num_units-1, -1, -1):
        top_decomp = "0"
        top_thickness = rel_addr(5,linenum)
        for cur_unit_i in range(top_unit_i,num_units):
            top_depth = "0"
            if cur_unit_i >= 1:
                top_depth = depth[cur_unit_i-1]
            if cur_unit_i > top_unit_i:
                top_decomp = rel_addr(7,linenum-1)
            total_sub_line = []
            total_sub_line.append("=" + unit[cur_unit_i])
            total_sub_line.append("=" + porosity_0[cur_unit_i])
            total_sub_line.append("={0}*EXP(-{1}*{2})".format(rel_addr(1,linenum), coeff[cur_unit_i], rel_addr(3,linenum)))
            total_sub_line.append("={0}+{1}".format(rel_addr(4,linenum), top_decomp))
            total_sub_line.append("=decomp({0},{1},{2},{3},{4})".format(porosity_0[cur_unit_i], coeff[cur_unit_i], depth[cur_unit_i], top_depth, top_decomp))
            total_sub_line.append("=2*" + rel_addr(4,linenum))
            total_sub_line.append("=(1-{0})*{1}/(1-{2})".format(porosity_p[cur_unit_i], thickness[cur_unit_i], rel_addr(2, linenum)))
            total_sub_line.append("=SUM({0}:{1})".format(top_thickness, rel_addr(5,linenum)))
            total_sub.append(total_sub_line)

            age_depth[cur_unit_i][top_unit_i+1] = "=" + rel_addr(7,linenum)
            if cur_unit_i == top_unit_i:
                dec_thickness[cur_unit_i] = rel_addr(6,linenum)
            if cur_unit_i == num_units-1:
                dec_depth[top_unit_i] = rel_addr(7,linenum)
            
            #cell formatting
            sheet.range(rel_addr(0,linenum) + ":" + rel_addr(7,linenum)).color = "#F2F2F2"
            sheet.range(rel_addr(0,linenum) + ":" + rel_addr(6,linenum)).api.Borders.Weight = 2
            sheet.range(rel_addr(7,linenum)).api.Borders.Weight = 3
            sheet.range(rel_addr(1,linenum) + ":" + rel_addr(7,linenum)).number_format = '0.00'
            if cur_unit_i < num_units-1:
                sheet.range(rel_addr(7,linenum)).api.Borders(9).Weight = 2
            if cur_unit_i > top_unit_i:
                sheet.range(rel_addr(7,linenum)).api.Borders(8).Weight = 2
            
            linenum = linenum + 1
        printProgressBar(num_units-top_unit_i, num_units, prefix = 'Total Subsidence:', suffix = 'Complete', length = 50)
        total_sub.append(["","","","","","","",""])
        linenum = linenum + 1
    #add block to sheet
    sheet.range((oy,ox)).value = total_sub

    print("Solving Subsidence (This may take a while)")

    #create tectonic subsidence block
    ox_sub = 0
    oy_sub = linenum

    ox = ox + ox_sub
    oy = oy + oy_sub
    _offset = (ox,oy)
    linenum = 0

    tect_sub = []
    tect_sub.append(["Tectonic Subsidence", "", "", "", "", "", "", "", "", ""])
    sheet.range(rel_addr(0,linenum) + ":" + rel_addr(1,linenum)).merge()
    sheet.range(rel_addr(0,linenum)).font.bold = True
    sheet.range(rel_addr(0,linenum)).color = "#FFFF00"
    linenum = linenum + 1
    
    tect_sub.append(["", "Thickness", "Porosity", "Dg", "Dw", "Ds", "Da", "Wd", "SL", "Z"])
    sheet.range(rel_addr(0,linenum) + ":" + rel_addr(9,linenum)).color = "#FFCCFF"
    sheet.range(rel_addr(0,linenum) + ":" + rel_addr(8,linenum)).api.Borders.Weight = 2
    sheet.range(rel_addr(9,linenum)).api.Borders.Weight = 3
    linenum = linenum + 1


    tect_sub_line = []
    for top_unit_i in range(num_units-1, -1, -1):
        top_thickness = rel_addr(1,linenum)
        top_density_s = rel_addr(5,linenum)
        for cur_unit_i in range(top_unit_i,num_units):
            tect_sub_line = []
            tect_sub_line.append("=" + unit[cur_unit_i])
            tect_sub_line.append("=" + rel_addr(5-ox_sub,linenum-oy_sub))
            tect_sub_line.append("=" + rel_addr(2-ox_sub,linenum-oy_sub))
            tect_sub_line.append("=" + density_g[cur_unit_i])
            tect_sub_line.append("=" + density_w[cur_unit_i])
            tect_sub_line.append("=(({0}*{1})+(1-{0})*{2})*{3}".format(rel_addr(2,linenum),rel_addr(4,linenum),rel_addr(3,linenum),rel_addr(1,linenum)))
            
            sheet.range(rel_addr(1,linenum) + ":" + rel_addr(9,linenum)).number_format = '0.00'
            sheet.range(rel_addr(0,linenum) + ":" + rel_addr(8,linenum)).color = "#FFFFCC"
            sheet.range(rel_addr(0,linenum) + ":" + rel_addr(8,linenum)).api.Borders.Weight = 2

            if cur_unit_i == top_unit_i:
                tect_sub_line.append("=" + density_a[cur_unit_i])
                tect_sub_line.append("=" + waterdepth[cur_unit_i])
                tect_sub_line.append("=" + sealevel[cur_unit_i])
                tect_sub_line.append("={0}*(({1}-{2})/({1}-{3}))+{4}-{5}*({1}/({1}-{3}))".format(rel_addr(1,linenum + num_units - top_unit_i),rel_addr(6,linenum),rel_addr(5,linenum + num_units - top_unit_i),rel_addr(4,linenum),rel_addr(7,linenum),rel_addr(8,linenum)))
                
                sheet.range(rel_addr(9,linenum)).color = "#FFFFCC"
                sheet.range(rel_addr(9,linenum)).api.Borders.Weight = 3

                tec_subsidence[cur_unit_i] = rel_addr(9,linenum)
            else:
                tect_sub_line = tect_sub_line + ["", "", "", ""]

            tect_sub.append(tect_sub_line)

            linenum = linenum + 1

        tect_sub.append(["","=SUM({0}:{1})".format(top_thickness, rel_addr(1,linenum-1)),"","","","=SUM({0}:{1})/{2}".format(top_density_s, rel_addr(5,linenum-1), rel_addr(1,linenum)),"","","",""])
        sheet.range(rel_addr(1,linenum) + ":" + rel_addr(9,linenum)).number_format = '0.00'
        printProgressBar(num_units-top_unit_i, num_units, prefix = 'Tect. Subsidence:', suffix = 'Complete', length = 50)
        linenum = linenum + 1

    sheet.range((oy,ox)).value = tect_sub

    #create depth block
    ox = tl[0]+17
    oy = tl[1]
    _offset = (ox,oy)

    for i in range(len(unit)):
        age_depth[i][0] = "=" + unit[i]
        age_depth[num_units][i+2] = "=" + age_bottom[i]
    age_depth[num_units][0] = "Age"
    age_depth[num_units][1] = "=" + age_top[0]
    
    sheet.range((oy, ox)).value = age_depth

    sheet.range(rel_addr(0, num_units) + ":" + rel_addr(num_units + 1, num_units)).color = "#F2F2F2"
    sheet.range(rel_addr(0, 0) + ":" + rel_addr(0, num_units)).color = "#F2F2F2"
    sheet.range(rel_addr(0, 0) + ":" + rel_addr(num_units + 1, num_units)).api.Borders.Weight = 2

    #create rate block
    ox = tl[0]+17
    oy = tl[1]+num_units+2
    _offset = (ox,oy)
    linenum = 0

    sub_rates = []
    sub_rates.append(["", "Ma", "Ma", "", "km", "km/my", "km", "km/my", "km", "km/my", "km", "km/my", "km", "km/my"])
    linenum = linenum + 1
    
    sub_rates.append(["", "Age Top", "mid-Age", "Age-Depth", "Thickness", "Sed Rate", "Dec. Thickness", "Dec. Sed. Rate", "Total (sed)", "Rate To. (sed)", "Total + WD", "Rate To. + WD", "Tectonic", "Rate Tec."])
    sheet.range(rel_addr(0,linenum) + ":" + rel_addr(3,linenum)).color = "#FFCC66"
    sheet.range(rel_addr(4,linenum) + ":" + rel_addr(5,linenum)).color = "#9BC2E6"
    sheet.range(rel_addr(6,linenum) + ":" + rel_addr(7,linenum)).color = "#DBDBDB"
    sheet.range(rel_addr(8,linenum) + ":" + rel_addr(9,linenum)).color = "#66FFFF"
    sheet.range(rel_addr(10,linenum) + ":" + rel_addr(11,linenum)).color = "#00CCFF"
    sheet.range(rel_addr(12,linenum) + ":" + rel_addr(13,linenum)).color = "#FFCCFF"
    sheet.range(rel_addr(0,linenum) + ":" + rel_addr(13,linenum)).api.Borders.Weight = 2
    linenum = linenum + 1

    sub_rates.append(["", "", "=" + age_top[0], "", "0", "0", "", "0", "", "0", "", "0", "", "0"])
    sheet.range(rel_addr(0,linenum) + ":" + rel_addr(13,linenum)).api.Borders.Weight = 2
    sheet.range(rel_addr(1,linenum) + ":" + rel_addr(13,linenum)).number_format = '0.00'
    linenum = linenum + 1

    for cur_unit_i in range(num_units):
        sub_rate_line = []
        sub_rate_line.append("=" + unit[cur_unit_i])
        sub_rate_line.append("=" + age_top[cur_unit_i])
        sub_rate_line.append("=({0}+{1})/2".format(rel_addr(1,linenum), rel_addr(1,linenum + 1)))
        sub_rate_line.append("={0}+{1}".format(rel_addr(4,linenum), rel_addr(3,linenum + 1)))
        sub_rate_line.append("=" + thickness[cur_unit_i])
        sub_rate_line.append("={0}/({1}-{2})".format(rel_addr(4,linenum),rel_addr(1,linenum + 1), rel_addr(1,linenum)))
        sub_rate_line.append("=" + dec_thickness[cur_unit_i])
        sub_rate_line.append("={0}/({1}-{2})".format(rel_addr(6,linenum),rel_addr(1,linenum + 1), rel_addr(1,linenum)))
        sub_rate_line.append("=" + dec_depth[cur_unit_i])
        sub_rate_line.append("=({0}-{1})/({2}-{3})".format(rel_addr(8,linenum),rel_addr(8,linenum + 1),rel_addr(1,linenum + 1),rel_addr(1,linenum)))
        sub_rate_line.append("={0}+{1}".format(rel_addr(8,linenum), waterdepth[cur_unit_i]))
        sub_rate_line.append("=({0}-{1})/({2}-{3})".format(rel_addr(10,linenum),rel_addr(10,linenum + 1),rel_addr(1,linenum + 1), rel_addr(1,linenum)))
        sub_rate_line.append("=" + tec_subsidence[cur_unit_i])
        sub_rate_line.append("=({0}-{1})/({2}-{3})".format(rel_addr(12,linenum),rel_addr(12,linenum + 1),rel_addr(1,linenum + 1), rel_addr(1,linenum)))

        sheet.range(rel_addr(0,linenum) + ":" + rel_addr(13,linenum)).api.Borders.Weight = 2
        sheet.range(rel_addr(1,linenum) + ":" + rel_addr(13,linenum)).number_format = '0.00'

        sub_rates.append(sub_rate_line)
        linenum = linenum + 1
        printProgressBar(cur_unit_i+1, num_units, prefix = 'Subsidence Rates:', suffix = 'Complete', length = 50)

    sub_rates.append(["", "=" + age_bottom[-1], "=" + rel_addr(1,linenum), "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"])
    sheet.range(rel_addr(0,linenum) + ":" + rel_addr(13,linenum)).api.Borders.Weight = 2
    sheet.range(rel_addr(1,linenum) + ":" + rel_addr(13,linenum)).number_format = '0.00'

    linenum = linenum + 1

    sheet.range((oy, ox)).value = sub_rates

    print ("Creating Charts")
    #Age-depth chart
    adc = sheet.charts.add()
    adc.left = sheet.range(rel_addr(0,0)).left
    adc.top = sheet.range(rel_addr(0,num_units+5)).top
    adc.height = sheet.range(rel_addr(0,num_units+22)).top - adc.top
    adc.width = sheet.range(rel_addr(9,0)).left - adc.left
    adc.chart_type = 'xy_scatter_lines'
    adc.api[1].SeriesCollection().NewSeries()
    adc.api[1].SeriesCollection(1).Name = 'Waterdepth'
    adc.api[1].SeriesCollection(1).XValues = sheet.range(rel_addr(2,3) + ":" + rel_addr(2,num_units+3)).api
    adc.api[1].SeriesCollection(1).Values = sheet.range(waterdepth[0] + ":" + waterdepth[-1]).api
    adc.api[1].SeriesCollection(1).MarkerStyle = -4142
    adc.api[1].SeriesCollection(1).Format.Line.ForeColor.RGB = rgb_to_int(hex_to_rgb("#00CCFF"))

    adc.api[1].SeriesCollection().NewSeries()
    adc.api[1].SeriesCollection(2).Name = 'Age-Depth Model'
    adc.api[1].SeriesCollection(2).XValues = sheet.range(rel_addr(1,3) + ":" + rel_addr(1,num_units+3)).api
    adc.api[1].SeriesCollection(2).Values = sheet.range(rel_addr(3,3) + ":" + rel_addr(3,num_units+3)).api
    adc.api[1].SeriesCollection(2).MarkerStyle = 1
    adc.api[1].SeriesCollection(2).Format.Line.ForeColor.RGB = rgb_to_int(hex_to_rgb("#FFC000"))
    adc.api[1].SeriesCollection(2).MarkerBackgroundColor = rgb_to_int(hex_to_rgb("#FFC000"))

    adc.api[1].Axes(1).ReversePlotOrder = True
    adc.api[1].Axes(1).HasMajorGridlines = True
    adc.api[1].Axes(1).HasTitle = True 
    adc.api[1].Axes(1).AxisTitle.Text = "Geologic Age (Ma)"

    adc.api[1].Axes(2).ReversePlotOrder = True
    adc.api[1].Axes(2).HasTitle = True 
    adc.api[1].Axes(2).AxisTitle.Text = "Depth (km)"

    #depth chart
    dc = sheet.charts.add()
    dc.left = sheet.range(rel_addr(9,0)).left
    dc.top = sheet.range(rel_addr(0,num_units+5)).top
    dc.height = sheet.range(rel_addr(0,num_units+22)).top - dc.top
    dc.width = sheet.range(rel_addr(18,0)).left - dc.left
    dc.chart_type = 'xy_scatter_lines'
    dc.api[1].SeriesCollection().NewSeries()
    dc.api[1].SeriesCollection(1).Name = 'Waterdepth'
    dc.api[1].SeriesCollection(1).XValues = sheet.range(rel_addr(2,3) + ":" + rel_addr(2,num_units+3)).api
    dc.api[1].SeriesCollection(1).Values = sheet.range(waterdepth[0] + ":" + waterdepth[-1]).api
    dc.api[1].SeriesCollection(1).MarkerStyle = -4142
    dc.api[1].SeriesCollection(1).Format.Line.ForeColor.RGB = rgb_to_int(hex_to_rgb("#00CCFF"))

    dc.api[1].SeriesCollection().NewSeries()
    dc.api[1].SeriesCollection(2).Name = sheet.range(rel_addr(8,1)).api
    dc.api[1].SeriesCollection(2).XValues = sheet.range(rel_addr(1,3) + ":" + rel_addr(1,num_units+3)).api
    dc.api[1].SeriesCollection(2).Values = sheet.range(rel_addr(8,3) + ":" + rel_addr(8,num_units+3)).api
    dc.api[1].SeriesCollection(2).MarkerStyle = 8
    dc.api[1].SeriesCollection(2).Format.Line.DashStyle = 4
    dc.api[1].SeriesCollection(2).Format.Line.ForeColor.RGB = rgb_to_int(hex_to_rgb("#00B050"))
    dc.api[1].SeriesCollection(2).MarkerBackgroundColorIndex = -4142
    dc.api[1].SeriesCollection(2).MarkerForegroundColor = rgb_to_int(hex_to_rgb("#00B050"))

    dc.api[1].SeriesCollection().NewSeries()
    dc.api[1].SeriesCollection(3).Name = sheet.range(rel_addr(10,1)).api
    dc.api[1].SeriesCollection(3).XValues = sheet.range(rel_addr(1,3) + ":" + rel_addr(1,num_units+3)).api
    dc.api[1].SeriesCollection(3).Values = sheet.range(rel_addr(10,3) + ":" + rel_addr(10,num_units+3)).api
    dc.api[1].SeriesCollection(3).Format.Line.DashStyle = 1
    dc.api[1].SeriesCollection(3).MarkerStyle = 8
    dc.api[1].SeriesCollection(3).Format.Line.ForeColor.RGB = rgb_to_int(hex_to_rgb("#00B050"))
    dc.api[1].SeriesCollection(3).MarkerBackgroundColorIndex = -4142

    dc.api[1].SeriesCollection().NewSeries()
    dc.api[1].SeriesCollection(4).Name = sheet.range(rel_addr(12,1)).api
    dc.api[1].SeriesCollection(4).XValues = sheet.range(rel_addr(1,3) + ":" + rel_addr(1,num_units+3)).api
    dc.api[1].SeriesCollection(4).Values = sheet.range(rel_addr(12,3) + ":" + rel_addr(12,num_units+3)).api
    dc.api[1].SeriesCollection(4).Format.Line.DashStyle = 1
    dc.api[1].SeriesCollection(4).MarkerStyle = 3
    dc.api[1].SeriesCollection(4).Format.Line.ForeColor.RGB = rgb_to_int(hex_to_rgb("#FF0000"))
    dc.api[1].SeriesCollection(4).MarkerBackgroundColor = rgb_to_int(hex_to_rgb("#FF0000"))


    dc.api[1].Axes(1).ReversePlotOrder = True
    dc.api[1].Axes(1).HasMajorGridlines = True
    dc.api[1].Axes(1).HasTitle = True 
    dc.api[1].Axes(1).AxisTitle.Text = "Geologic Age (Ma)"

    dc.api[1].Axes(2).ReversePlotOrder = True
    dc.api[1].Axes(2).HasTitle = True 
    dc.api[1].Axes(2).AxisTitle.Text = "Depth (km)"


    #subsidence rate chart
    subrc = sheet.charts.add()
    subrc.left = sheet.range(rel_addr(9,0)).left
    subrc.top = sheet.range(rel_addr(0,num_units+22)).top
    subrc.height = sheet.range(rel_addr(0,num_units+32)).top - subrc.top
    subrc.width = sheet.range(rel_addr(18,0)).left - subrc.left
    subrc.chart_type = 'xy_scatter_lines'
    subrc.api[1].SeriesCollection().NewSeries()
    subrc.api[1].SeriesCollection(1).Name = sheet.range(rel_addr(11,1)).api
    subrc.api[1].SeriesCollection(1).XValues = sheet.range(rel_addr(2,2) + ":" + rel_addr(2,num_units+3)).api
    subrc.api[1].SeriesCollection(1).Values = sheet.range(rel_addr(11,2) + ":" + rel_addr(11,num_units+3)).api
    subrc.api[1].SeriesCollection(1).Format.Line.DashStyle = 1
    subrc.api[1].SeriesCollection(1).MarkerStyle = 8
    subrc.api[1].SeriesCollection(1).Format.Line.ForeColor.RGB = rgb_to_int(hex_to_rgb("#00B050"))
    subrc.api[1].SeriesCollection(1).MarkerBackgroundColorIndex = -4142

    subrc.api[1].SeriesCollection().NewSeries()
    subrc.api[1].SeriesCollection(2).Name = sheet.range(rel_addr(9,1)).api
    subrc.api[1].SeriesCollection(2).XValues = sheet.range(rel_addr(2,2) + ":" + rel_addr(2,num_units+3)).api
    subrc.api[1].SeriesCollection(2).Values = sheet.range(rel_addr(9,2) + ":" + rel_addr(9,num_units+3)).api
    subrc.api[1].SeriesCollection(2).MarkerStyle = 8
    subrc.api[1].SeriesCollection(2).Format.Line.DashStyle = 4
    subrc.api[1].SeriesCollection(2).Format.Line.ForeColor.RGB = rgb_to_int(hex_to_rgb("#00B050"))
    subrc.api[1].SeriesCollection(2).MarkerBackgroundColorIndex = -4142

    subrc.api[1].SeriesCollection().NewSeries()
    subrc.api[1].SeriesCollection(3).Name = sheet.range(rel_addr(13,1)).api
    subrc.api[1].SeriesCollection(3).XValues = sheet.range(rel_addr(2,2) + ":" + rel_addr(2,num_units+3)).api
    subrc.api[1].SeriesCollection(3).Values = sheet.range(rel_addr(13,2) + ":" + rel_addr(13,num_units+3)).api
    subrc.api[1].SeriesCollection(3).Format.Line.DashStyle = 1
    subrc.api[1].SeriesCollection(3).MarkerStyle = 3
    subrc.api[1].SeriesCollection(3).Format.Line.ForeColor.RGB = rgb_to_int(hex_to_rgb("#FF0000"))
    subrc.api[1].SeriesCollection(3).MarkerBackgroundColor = rgb_to_int(hex_to_rgb("#FF0000"))


    subrc.api[1].Axes(1).ReversePlotOrder = True
    subrc.api[1].Axes(1).HasMajorGridlines = True
    subrc.api[1].Axes(1).HasTitle = True 
    subrc.api[1].Axes(1).AxisTitle.Text = "Geologic Age (Ma)"

    subrc.api[1].Axes(2).HasTitle = True 
    subrc.api[1].Axes(2).AxisTitle.Text = "Subsidence Rate (km/My)"

    #sedimentation rate chart
    sedrc = sheet.charts.add()
    sedrc.left = sheet.range(rel_addr(0,0)).left
    sedrc.top = sheet.range(rel_addr(0,num_units+22)).top
    sedrc.height = sheet.range(rel_addr(0,num_units+32)).top - sedrc.top
    sedrc.width = sheet.range(rel_addr(9,0)).left - sedrc.left
    sedrc.chart_type = 'xy_scatter_lines'
    sedrc.api[1].SeriesCollection().NewSeries()
    sedrc.api[1].SeriesCollection(1).Name = sheet.range(rel_addr(5,1)).api
    sedrc.api[1].SeriesCollection(1).XValues = sheet.range(rel_addr(2,2) + ":" + rel_addr(2,num_units+3)).api
    sedrc.api[1].SeriesCollection(1).Values = sheet.range(rel_addr(5,2) + ":" + rel_addr(5,num_units+3)).api
    sedrc.api[1].SeriesCollection(1).Format.Line.DashStyle = 1
    sedrc.api[1].SeriesCollection(1).MarkerStyle = 1
    sedrc.api[1].SeriesCollection(1).Format.Line.ForeColor.RGB = rgb_to_int(hex_to_rgb("#FFC000"))
    sedrc.api[1].SeriesCollection(1).MarkerBackgroundColor = rgb_to_int(hex_to_rgb("#FFC000"))


    sedrc.api[1].SeriesCollection().NewSeries()
    sedrc.api[1].SeriesCollection(2).Name = sheet.range(rel_addr(7,1)).api
    sedrc.api[1].SeriesCollection(2).XValues = sheet.range(rel_addr(2,2) + ":" + rel_addr(2,num_units+3)).api
    sedrc.api[1].SeriesCollection(2).Values = sheet.range(rel_addr(7,2) + ":" + rel_addr(7,num_units+3)).api
    sedrc.api[1].SeriesCollection(2).MarkerStyle = 1
    sedrc.api[1].SeriesCollection(2).Format.Line.ForeColor.RGB = rgb_to_int(hex_to_rgb("#C55A11"))
    sedrc.api[1].SeriesCollection(2).MarkerBackgroundColor = rgb_to_int(hex_to_rgb("#C55A11"))

    sedrc.api[1].Axes(1).ReversePlotOrder = True
    sedrc.api[1].Axes(1).HasMajorGridlines = True
    sedrc.api[1].Axes(1).HasTitle = True 
    sedrc.api[1].Axes(1).AxisTitle.Text = "Geologic Age (Ma)"

    sedrc.api[1].Axes(2).HasTitle = True 
    sedrc.api[1].Axes(2).AxisTitle.Text = "Sedimentation Rate (km/My)"
    print ("Done")

'''
@xw.sub
def charts():
    #create rate block
    ox = 33
    oy = 14
    global _offset
    _offset = (ox,oy)

    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    num_units=6
    global _sheet
    _sheet = sheet'''

@xw.func
@xw.arg('phi_0', doc='Initial porosity')
@xw.arg('c', doc='Coefficient')
@xw.arg('top_p', doc='Present top depth')
@xw.arg('bottom_p', doc='Present bottom depth')
@xw.arg('top_decomp', doc='Decompacted top depth')
def decomp( phi_0, c, top_p, bottom_p, top_decomp):
    """Performs decompaction based on the following parameters:
    (Initial porosity, Coefficient c, Present top depth, Present bottom depth, Decompacted top depth)"""
    thickness_p = bottom_p - top_p
    if thickness_p == 0:
        return 0
    center_p = (bottom_p + top_p)/2
    phi_p = phi_0 * math.exp(-center_p*c)
    d = Symbol('d')
    f = (1-phi_p) * thickness_p / (1-(phi_0 * exp(-(d + top_decomp)*c))) + (2*d)
    return nsolve(f,d,0)