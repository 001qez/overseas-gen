#!/usr/bin/env python
# -*- coding: utf-8 -*-

# import LatLon, MSO
# author = LJS
# date = 16 Feb 2016
# version = 0.1

import numpy as np
import matplotlib.pyplot as plt


################# modify here ##########
import sys, os
if getattr(sys, 'frozen', False):
    os.environ['BASEMAPDATA'] = os.path.join(os.path.dirname(sys.executable), 'mpl-basemap-data')
from mpl_toolkits.basemap import Basemap
import FileDialog, pyproj
########################################

import LatLon, math, csv

import win32com.client, os.path, sys, io, json, collections
import MSO, MSPPT
g = globals()
for c in dir(MSO.constants):    g[c] = getattr(MSO.constants, c)
for c in dir(MSPPT.constants):  g[c] = getattr(MSPPT.constants, c)



def roundup(x, num):
    return int(math.ceil(x / num)) * num
def rounddown(x, num):
    return int(math.floor(x / num)) * num


def alt_create_map(sw_corner, ne_corner, filename):
    
    distance = sw_corner.distance(ne_corner)
    #print "coverage area diagonal distance is", distance, "km"
    
    # determine the scale of the map area and 
    # interval for meridian and parallel
    if distance < 460:
        scale = 1.50
        interval = 2
    elif distance < 1100:
        scale = 1.25
        interval = 2
    elif distance < 2250:
        scale = 0.75
        interval = 5
    else:
        scale = 0.50
        interval = 10
    
    # to avoid the map being too wide or too tall
    se_corner = LatLon.LatLon(sw_corner.lat.decimal_degree, ne_corner.lon.decimal_degree)
    lat_distance = se_corner.distance(ne_corner)
    lon_distance = se_corner.distance(sw_corner)
    if 4 * lat_distance > 5 * lon_distance:
        lon_distance = 1.25 * lat_distance
    if 4 * lon_distance > 5 * lat_distance:
        lat_distance = 0.8 * lon_distance
    
    # scale the map area
    lat_distance = scale * lat_distance
    lon_distance = scale * lon_distance
    
    map_sw_corner = sw_corner.offset(180, lat_distance).offset(270, lon_distance)
    map_ne_corner = ne_corner.offset(0, lat_distance).offset(90, lon_distance)
    
    m = Basemap(llcrnrlat=map_sw_corner.lat.decimal_degree, 
                llcrnrlon=map_sw_corner.lon.decimal_degree, 
                urcrnrlat=map_ne_corner.lat.decimal_degree, 
                urcrnrlon=map_ne_corner.lon.decimal_degree, 
                projection='merc', resolution='i')
    
    m.fillcontinents(color='0.99',lake_color='white')
    m.drawcoastlines(linewidth=0.5)
    
    # draw meridians and parallels with the interval chosen
##    meridians = np.arange(rounddown(m.llcrnrlon,interval), roundup(m.urcrnrlon,interval), interval)
##    m.drawmeridians(meridians, color='0.75', dashes=[1, 1], labels=[0,0,0,1], 
##                    fontsize=8.5, yoffset = -0.022*(m.ymax-m.ymin))
##    parallels = np.arange(rounddown(m.llcrnrlat,interval), roundup(m.urcrnrlat,interval), interval)
##    m.drawparallels(parallels, color='0.75', dashes=[1, 1], labels=[1, 0, 0, 0], 
##                    fontsize=8.5, xoffset = -0.044*(m.xmax-m.xmin))
    meridians = np.arange(rounddown(m.llcrnrlon,interval), roundup(m.urcrnrlon,interval), interval)
    m.drawmeridians(meridians, color='0.75', dashes=[1, 1], labels=[0,0,0,1], 
                    fontsize=8)
    parallels = np.arange(rounddown(m.llcrnrlat,interval), roundup(m.urcrnrlat,interval), interval)
    m.drawparallels(parallels, color='0.75', dashes=[1, 1], labels=[1, 0, 0, 0], 
                    fontsize=8)
    #aaxx = plt.gca()
    #aaxx.axis[:].invert_ticklabel_direction()
    
    # Draw area of interest
    lats = [sw_corner.lat.decimal_degree, 
            sw_corner.lat.decimal_degree, 
            ne_corner.lat.decimal_degree, 
            ne_corner.lat.decimal_degree, 
            sw_corner.lat.decimal_degree]
    lons = [ne_corner.lon.decimal_degree, 
            sw_corner.lon.decimal_degree, 
            sw_corner.lon.decimal_degree, 
            ne_corner.lon.decimal_degree, 
            ne_corner.lon.decimal_degree]
    x, y = m(lons, lats)
    m.plot(x, y, ':', linewidth=2, color='k', dash_capstyle='round') 
    
    fig = plt.gcf()
    fig.set_size_inches(fig.get_size_inches()*0.75)
    
    plt.savefig(filename, bbox_inches='tight', pad_inches=0.05)
    plt.close()

def create_ppt(row):

    Application = win32com.client.Dispatch("PowerPoint.Application")
    
    Presentation = Application.Presentations.Open(
        os.path.join(os.getcwd(), row['type'] + '-template.ppt'), False, True, False)
    print os.path.join(os.getcwd(), row['type'] + '.ppt')
    slide = Presentation.Slides(1)
    
    try:
        print row['main_title']
        # change the main_title
        table_0_shape = slide.Shapes('TABLE_0')
        table_0_shape.Table.Rows(1).Cells.Item(1).Shape.TextFrame.TextRange.Text = row['main_title']
        table_1_shape = slide.Shapes('TABLE_1')
        table_1_shape.Table.Rows(1).Cells.Item(1).Shape.TextFrame.TextRange.Text = row['main_title']
        table_2_shape = slide.Shapes('TABLE_2')
        table_2_shape.Table.Rows(1).Cells.Item(1).Shape.TextFrame.TextRange.Text = row['main_title']
    except:
        print 'Error for main_title'
    
    try:
        print row['valid']
        # change the valid
        table_0_shape.Table.Rows(2).Cells.Item(1).Shape.TextFrame.TextRange.Text = 'Valid: ' + row['valid']
    except:
        print 'Error for valid'
    
    try:
        print row['issue']
        # change the issue
        table_0_shape.Table.Rows(2).Cells.Item(2).Shape.TextFrame.TextRange.Text = 'Issued: ' + row['issue']
    except:
        print 'Error for issue'
    
    # change the area of interest
    try:
        sw_corner = LatLon.string2latlon(row['sw lat'], row['sw lon'], 'd% %m% %H')
        ne_corner = LatLon.string2latlon(row['ne lat'], row['ne lon'], 'd% %m% %H')
    except:
        print 'Error with sw lat, sw lon, ne lat, or ne lon'
    
    def pLAT(L):
        return "".join((str(abs(int(L.lat.to_string('d')))).zfill(2),
                        unichr(176), 
                        str(abs(int(L.lat.to_string('%m')))).zfill(2), 
                        "'", L.lat.to_string('%H')))
    def pLON(L):
        return "".join((str(abs(int(L.lon.to_string('d')))).zfill(3),
                        unichr(176), 
                        str(abs(int(L.lon.to_string('%m')))).zfill(2), 
                        "'", L.lon.to_string('%H')))
    try:
        temp = "".join((pLAT(sw_corner), ' ', '-', ' ',
                        pLAT(ne_corner), '; ', 
                        pLON(sw_corner), ' ', '-', ' ', 
                        pLON(ne_corner)))
        print temp
        table_1_shape = slide.Shapes('TABLE_1')
        table_1_shape.Table.Rows(2).Cells.Item(2).Shape.TextFrame.TextRange.Text = temp
    except:
        print 'Error inserting area of coverage'
    
    try:
        alt_create_map(sw_corner, ne_corner, row['file_code']+'.png')
    except:
        print 'Error with file_code or creating map image'
    
    try:
        # insert image into ppt
        shp = slide.Shapes.AddPicture(
            os.path.join(os.getcwd(), row['file_code']+'.png'),
            False, True, 50, 80)
        shp.Name = 'map'
        shp.ZOrder(msoSendBackward)
    except:
        print 'Error inserting map image into ppt'
    
    try:
        temp = row['file_code'] + '_' + row['type'] + '_' + row['DisplayDays'] + '_' + row['AreaNameKey'] + '.ppt'
        Presentation.SaveCopyAs(os.path.join(os.getcwd(), temp))
        print os.path.join(os.getcwd(),temp)
    except:
        print 'Error with DisplayDays, AreaNameKey. Error saving the powerpoint file.'
    
    try:
        # close the powerpoint file
        Presentation.Close()
        Application.Quit()
        print '\n'
    except:
        print 'Error closing the powerpoint file'

##def create_ppt(row):
##    
##    try:
##        print row['type']
##        Application = win32com.client.Dispatch("PowerPoint.Application")
##        Presentation = Application.Presentations.Open(os.path.join( os.getcwd(),
##                                                                    row['type']+'.ppt'))
##        # powerpoint = open(row['type']+'.ppt')
##    except:
##        print 'Error for type or opening template ppt'
##    
##    try:
##        slide = Presentation.Slides(1)
##        slide.Shapes('TABLE_0').Table.Rows(1).Cells.Item(1).Shape.TextFrame.TextRange.Text = row['main_title']
##        print row['main_title']
##        # change the main_title
##    except:
##        print 'Error for main_title'
##    
##    try:
##        print row['valid']
##        # change the valid
##    except:
##        print 'Error for valid'
##    
##    try:
##        print row['issue']
##        # change the issue
##    except:
##        print 'Error for issue'
##    
##    # change the area of interest
##    try:
##        sw_corner = LatLon.string2latlon(row['sw lat'], row['sw lon'], 'd% %m% %H')
##        ne_corner = LatLon.string2latlon(row['ne lat'], row['ne lon'], 'd% %m% %H')
##    except:
##        print 'Error with sw lat, sw lon, ne lat, or ne lon'
##    
##    def pLAT(L):
##        return "".join((str(abs(int(L.lat.to_string('d')))).zfill(2),
##                        u'\N{DEGREE SIGN}', 
##                        str(abs(int(L.lat.to_string('%m')))).zfill(2), 
##                        "'", L.lat.to_string('%H')))
##    def pLON(L):
##        return "".join((str(abs(int(L.lon.to_string('d')))).zfill(3),
##                        u'\N{DEGREE SIGN}', 
##                        str(abs(int(L.lon.to_string('%m')))).zfill(2), 
##                        "'", L.lon.to_string('%H')))
##    try:
##        print "".join((pLAT(sw_corner), ' ', u'\N{EN DASH}', ' ', 
##                       pLAT(ne_corner), '; ', 
##                       pLON(sw_corner), ' ', u'\N{EN DASH}', ' ', 
##                       pLON(ne_corner)))
##    except:
##        print 'Error inserting area of coverage'
##    
##    try:
##        alt_create_map(sw_corner, ne_corner, row['file_code']+'.png')
##    except:
##        print 'Error with file_code or creating map image'
##    
##    try:
##        # insert image into ppt
##        pass
##    except:
##        print 'Error inserting map image into ppt'
##    
##    try:
##        Presentation.SaveAs(os.path.join(os.getcwd(), row['file_code'] + '_' + row['type'] + '_' + row['DisplayDays'] + '_' + row['AreaNameKey'] + '.ppt'))
##        print row['file_code'] + '_' + row['type'] + '_' + row['DisplayDays'] + '_' + row['AreaNameKey'] + '.ppt'
##    except:
##        print 'Error with DisplayDays, AreaNameKey. Error saving the powerpoint file.'
##    
##    try:
##        # close the powerpoint file
##        Presentation.Close()
##    except:
##        print 'Error closing the powerpoint file'
##
##    Application.Quit()

print 'If you have any open PowerPoint, close them before proceeding...'
raw_input('Press Enter to proceed')

with open('input.csv') as csvfile:
    
    try:
        reader = csv.DictReader(csvfile)
    except:
        print "input.csv could not be read or found"

    # Application = win32com.client.Dispatch("PowerPoint.Application")
    
    for row in reader:
        try:
            create_ppt(row)
        except:
            print "could not create ppt"

    # Application.Quit()

raw_input('Press Enter to close')
