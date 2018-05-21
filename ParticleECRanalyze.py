# Merges .ecr and .txt to get time-correlated data and creates deformation/resistance(time) plot in excel
# Output: Excel spreadsheet containing the following columns: Time, Force, Disp, Stress, Strain, Voltage, Current, Resistance + some plots
# Empty resistance values where the resistance is noise (e.g. negative or extremely high values)

# Intended for the analysis of spherical particles (strain calculations, etc are based on spheres)

import sys
import getopt
import operator
import os.path
from glob import glob
from Tkinter import *
import tkFileDialog
import xlsxwriter
import math
import numpy as np
from measurement import measurement
from mbox import mbox

IN_FOLDER = '/'

thresholds = [5, 10, 100] # Resistance thresholds under which to calculate strain
strains = [0.1, 0.15, 0.2, 0.3, 0.35, 0.4, 0.45, 0.5, 0.55, 0.6] # Strains at which to find resistance
  
def writeToXlsx(filename, data):
    split_file = filename.split('/')
    savename = split_file[-1]
    # print(savename)
    name = savename+'.xlsx'
    workbook = xlsxwriter.Workbook(filename+'.xlsx')
    
    # Initialize plots for all measurement sets
    chartsheet = workbook.add_worksheet('Analysis')
    r_stress = workbook.add_chart({'type':'scatter'})
    r_stress.set_x_axis({
        'name':'Nominal stress',
        'min':0    
    })
    r_stress.set_y_axis({
        'name':'Resistance [Ohm]',
        'major_gridlines':{'visible':False},
        'min':0 
    })
    
    r_strain = workbook.add_chart({'type':'scatter'})
    r_strain.set_x_axis({
        'name':'Nominal strain',
        'min':0    
    })
    r_strain.set_y_axis({
        'name':'Resistance [Ohm]',
        'major_gridlines':{'visible':False},
        'min':0
    })
    
    stress_strain = workbook.add_chart({'type':'scatter'})
    stress_strain.set_x_axis({
        'name':'Nominal strain',
        'min':0    
    })
    stress_strain.set_y_axis({
        'name':'Nominal stress [MPa]',
        'major_gridlines':{'visible':False},
        'min':0
    })
    
    # STATISTICS SHEET
    statSheet = workbook.add_worksheet('Statistics')
    statSheet.write(0,0, 'Data series')
    statHeaders = []
    statCounter = 1 # Counter contains the index of the next column to be filled in the stat sheet.
    sweepNumber = None
    sweepPlotInitialized = False
    
    for idx, m in enumerate(data):
        
        if m.sweepFound and not sweepPlotInitialized:
            sweepNumber = 0
            sweepSheet = workbook.add_worksheet('Sweep compare')
            sweep_compare = workbook.add_chart({'type':'scatter'})
            sweep_compare.set_x_axis({
                'name':'I [A]'
            })
            sweep_compare.set_y_axis({
                'name':'V [V]',
                'major_gridlines':{'visible':False}
            })
            sweep_compare.set_title({'name': 'I-V sweeps where applicable'})
            sweepPlotInitialized = True
            
        # New sheet for each measurment
        worksheet = workbook.add_worksheet(m.fileName)
        num_datapoints = str(len(m.data))
        # Write header data
        header = ['Time [s]','Depth [nm]','Force [uN]','Strain', 'Stress [MPa]', 'Current[A]', 'Voltage [V]','Resistance [Ohm]']
        row = 0
        col = 0
        for item in header:
            worksheet.write(row, col, item)
            col += 1  
        # Write measurement data to sheet
        row = 0
        for dataRow in m.data:
            col = -1
            row += 1
            for item in dataRow:
                col += 1
                try:
                    worksheet.write(row,col,float(item))
                except ValueError:
                    # Value error when empty data field, skip
                    continue
            
        # First chart: resistance (y1) and deformation (y2) as a function of indentation time
        chart = workbook.add_chart({'type':'scatter'})
        # Resistance series
        chart.add_series({
            'name':'='+m.fileName+'!$H$1',
            'categories':'='+m.fileName+'!$A$2:$A$'+num_datapoints,
            'values':'='+m.fileName+'!$H$2:$H$'+num_datapoints
        })
        chart.set_y_axis({
            'min':0,
            'major_gridlines':{'visible':False},
            'name':'Resistance [Ohm]'
        })
        #Deformation series
        chart.add_series({
            'name':'='+m.fileName+'!$B$1',
            'categories':'='+m.fileName+'!$A$2:$A$'+num_datapoints,
            'values':'='+m.fileName+'!$B$2:$B$'+num_datapoints,
            'y2_axis':True
        })
        chart.set_y2_axis({
            'min':0,
            'name':'Deformation [nm]'
        })
        chart.set_x_axis({
            'name':'Time [s]'
        })
        worksheet.insert_chart('J2', chart)
    
        #Second chart: Force vs deformation
        chart2 = workbook.add_chart({'type':'scatter'})
        chart2.add_series({
            'name':'='+m.fileName+'!$C$1',
            'categories':'='+m.fileName+'!$B$2:$B$'+num_datapoints,
            'values':'='+m.fileName+'!$C$2:$C$'+num_datapoints
        })
        chart2.set_y_axis({
            'min':0,
            'name':'Force[uN]',
            'major_gridlines':{'visible':False}
        })
        chart2.set_x_axis({
            'min':0,
            'name':'Deformation [nm]'
        })
        chart2.set_title({'none':True})
        chart2.set_legend({'none':True})
        worksheet.insert_chart('J17', chart2)
    
        # Third chart : Resistance vs deformation       
        chart3 = workbook.add_chart({'type':'scatter'})
        chart3.add_series({
            'categories':'='+m.fileName+'!$E$2:$E$'+str(m.maxI),
            'values':'='+m.fileName+'!$H$2:$H$'+str(m.maxI)
        })
        chart3.set_y_axis({
            'min':0,
            'name':'Resistance [Ohm]',
            'major_gridlines':{'visible':False}
        })
        chart3.set_x_axis({
            'min':0,
            'name':'Nominal stress [MPa]'
        })
        chart3.set_title({'none':True})
        chart3.set_legend({'none':True})
        worksheet.insert_chart('J35', chart3)
        
        # Add series to the comparative charts
        r_stress.add_series({
            'categories':'='+m.fileName+'!$E$2:$E$'+str(m.maxI), # Categories are x-values
            'values':'='+m.fileName+'!$H$2:$H$'+str(m.maxI), # Values are y-values
            'name':m.fileName
        })   
        r_strain.add_series({
            'categories':'='+m.fileName+'!$B$2:$B$'+str(m.maxI),
            'values':'='+m.fileName+'!$H$2:$H$'+str(m.maxI),
            'name':m.fileName
        })        
        stress_strain.add_series({
            'categories':'='+m.fileName+'!$D$2:$D$'+str(m.maxI),
            'values':'='+m.fileName+'!$E$2:$E$'+str(m.maxI),
            'name':m.fileName
        }) 
        
        # If applicable, add I-V sweep to comparative chart
        if m.sweepFound:
            # Write I-V data
            sweepSheet.write(0, sweepNumber, m.fileName + ', intercept: ')
            sweepSheet.write(0, sweepNumber+1, m.sweepFit[1]) # Writes intercept to chart
            sweepSheet.write(1, sweepNumber, 'I')
            sweepSheet.write(1, sweepNumber+1, 'V')
            #print(m.sweepData)
            for i in range(len(m.sweepData)):
                sweepSheet.write(2+i, sweepNumber, m.sweepData[i][0])
                sweepSheet.write(2+i, sweepNumber+1, m.sweepData[i][1])
            alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
            sweep_compare.add_series({
                'categories':'=\'Sweep compare\'!'+alphabet[sweepNumber]+'3:'+alphabet[sweepNumber]+str(len(m.sweepData)+2),
                'values':'=\'Sweep compare\'!'+alphabet[sweepNumber+1]+'3:'+alphabet[sweepNumber+1]+str(len(m.sweepData)+2),
                'name':m.fileName
            })
            sweepNumber += 2 
            
  
        statSheet.write(idx+1,0,m.fileName)
        # Iterate over statistics variable and add missing headers and data to sheet
        for key, value in m.statistics.iteritems():
            found = False
            for r, header in enumerate(statHeaders):
                if header == key:
                    statSheet.write(idx+1,r+1,value)
                    found = True
                    break       
            if found == False: # Then we need to create a new column to fill out
                statHeaders.append(key)
                statSheet.write(0,statCounter,key)
                statSheet.write(idx+1,statCounter,value)
                statCounter += 1

    statSheet.write(idx+2,0,'Average')
    statSheet.write(idx+3,0,'Stdev')
    statSheet.write(idx+4,0,'\% Particles no data')
    
    ab = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    # Calculate stddev and averages of all statistical quantities
    for r, header in enumerate(statHeaders):
        statSheet.write_formula(idx+2,r+1,'=AVERAGE('+str(ab[r])+'2:'+str(ab[r])+str(idx+2)+')')
        statSheet.write_formula(idx+3,r+1,'=STDEV('+str(ab[r])+'2:'+str(ab[r])+str(idx+2)+')')
        statSheet.write_formula(idx+4,r+1,'=COUNTBLANK('+str(ab[r])+'2:'+str(ab[r])+str(idx+2)+')*100/'+str(idx+1))
        
    chartsheet.insert_chart('A1', r_stress)
    chartsheet.insert_chart('J1', r_strain)
    chartsheet.insert_chart('A20', stress_strain)
    if sweepPlotInitialized:
        chartsheet.insert_chart('J20', sweep_compare)
        
    workbook.close()
    return 
    
def usage():
    print('This script merges .ecr and .txt to get time-correlated data and creates deformation/resistance(time) plot in excel.\n Usage: python ECRmulti_v2.py (-u -c ) \n Options: -u, --unloading: include loading data in plot. Default is to plot loading data only. \n -c --clean: removes mechanical data points that don\'t have electrical data associated with them. WARNING: will result in incomplete mechanical data.')

if __name__ == '__main__':
    currentpath = os.path.dirname(os.path.abspath(__file__))
    pathfile = currentpath + '/dataPath.txt'
    print(currentpath)
    if os.path.isfile(pathfile):
        # print('Meep')
        f = open(pathfile, 'r')
        directory = f.read()
        f.close()
    else:
        directory = IN_FOLDER
        
         
    UNLOADING = False # Toggle plot with/without unloading data for analysis
    CLEANED = False # Toggle removal of mechanical data points for which there is no corresponding electrical data
    
    '''
    try:
        opts, args = getopt.getopt(sys.argv[1:], "huc", ["help", "unloading", "clean"])
        # print opts, args
    except getopt.GetoptError:          
            usage()                         
            sys.exit(2)                     
    for opt, arg in opts:
        if opt in ("-h", "--help"):
            usage()                     
            sys.exit()                  
        elif opt in ("-u", "--unloading"):
            UNLOADING = True
        elif opt in ("-c", "--clean"):
            CLEANED = True
    '''
    
    root = Tk()
    root.lift()
    inFiles = tkFileDialog.askopenfilename(parent=root, title='Select .txt files', multiple=True, filetypes=[('text files', '*.txt')], initialdir=directory)
    if not inFiles:
        sys.exit('Program cancelled')
    path = inFiles[0]
    f = open(pathfile, 'w')
    path = inFiles[0]
    f.write(path)
    f.close()
    
    #Specify name you want to save to
    filename = tkFileDialog.asksaveasfilename(parent=root, title='Save file with multiple plots as:', initialdir=path)
    if not filename:
        sys.exit('Program cancelled')
    # Specify particle size (Optional??)
    size = mbox('Enter the particle diameter in um:  ', entry=True)
    
    data = []
    for inFile in inFiles:
        print(inFile)
        current = measurement(inFile, float(size))
        if CLEANED == True:       
            current.clean()
        for item in thresholds:
            current.findThresholdStrain(item)
        for item in strains:
            current.findResistanceAtStrain(item)
        data.append(current)
    writeToXlsx(filename, data)