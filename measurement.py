# Class for coupled ECR/mechanical nanoindenter measurements

from glob import glob
import math
import numpy as np
from numpy import linalg
import re
from collections import namedtuple
from collections import OrderedDict

mintimediff = 0.02 # For data point time correlation

class DataPoint:
   
   def __init__(self, t, f, d, stress, strain, I='', V='', R=''):
       self.t=t
       self.f=f
       self.d=d
       self.stress=stress
       self.strain=strain
       self.I=I
       self.V=V
       self.R=R
       return   
   
   def __iter__(self):
       for item in  (self.t, self.d, self.f, self.strain, self.stress, self.I, self.V, self.R):
           yield item

class measurement:
# Initialization: time-correlates electrical and mechanical data and creates a set of arrays
    def __init__(self, fileName, size=0):
        self.particleSize = size
        split_file = fileName.split('/')
        saveName = split_file[-1]
        self.fileName = saveName # only 'A LC.txt'
        self.filePath = fileName # Full path to location
        self.data = []
        if isempty(self.data.r):
            self.contact = False
        else:
            self.contact = True
        self.statistics = OrderedDict() # For finding at what strain the data is below a certain resistance
        self.sweepStartTime = None
        self.sweepEndTime = None
        self.sweepFound = False
        self.sweepFit = None
        self.sweepData = []
        self.merge()
        if self.sweepFound:
            self.fitSweep()
        self.setMinR()
        self.setMaxI()
        self.setRecoveryRatio()
        return
        
    def merge(self):
        read = False
        strippedfile = self.filePath[:-4] # Removes .txt from end of file name
        t = []
        f = []
        d = []
        stress = []
        # First parse mechanical data text file (which has the most data)
        i = 0
        if self.particleSize > 0:
            area = np.power((1e-6*self.particleSize/2), 2)*math.pi # m^2
            length = self.particleSize*1e3 # Units: nm
        with open(self.filePath, 'U') as data:
            for line in data:
                parts = line.split("\t")
                if read == True:
                    try:
                        (t, f, d) = (float(parts[2]), float(parts[1]), float(parts[0]))
                    except IndexError:
                        # Empty row, skip
                        continue
                    stress = None
                    strain = None
                    if self.particleSize > 0:
                        stress = (float(parts[1])*1e-12)/area # Units: MPa
                        strain = float(parts[0])/length
                    self.data.append(DataPoint(t, f, d, stress, strain))
                if parts[0] == "Depth (nm)":
                    read = True
        # Then parse ECR file and match time stamps
            ECR = strippedfile+'.ecr' 
            read = False
            currentval = 0 # index
            recordsweep = False
        try:
            with open(ECR, 'U') as data:
                for line in data:
                    temp = line.split("\t")
                    parts = list(filter(None, temp)) # Removes empty entries
                    #print(parts)
                    if parts[0] == 'Voltage(V) ': # Start reading from the beginning of the sweep
                        read = True
                        continue  
                    if read and len(parts) > 1:
                        time = float(parts[2].strip('\n'))
                        #print(time)
                        if self.sweepFound and time == self.sweepStartTime:
                            print('Start recording sweep!')
                            recordsweep = True
                        if self.sweepFound and recordsweep and time > self.sweepEndTime:
                            print('Stopped recording sweep at ' + str(time))
                            recordsweep = False # turn off recording
                        bestfit = mintimediff
                        for i in range(currentval, len(self.data)-1):
                            timediff = abs(self.data[i].t-time)
                            if timediff < mintimediff: # Check that time points are in the correct range
                                if bestfit > timediff: # Then this fit is the best so far
                                    bestfit = timediff
                                    continue
                                if bestfit < timediff: # Then we have passed the point of best fit
                                    currentval = i # Reset so we start from this point when fitting the next data point
                                    (I, V) = (float(parts[1]), float(parts[0]))
                                    self.data[i-1].V = V
                                    self.data[i-1].I = I
                                    if recordsweep and I != 0:
                                        self.sweepData.append([I,V])
                                    try:
                                        r = V/I
                                    except ZeroDivisionError: 
                                        continue # Don't need to redefine the resistance
                                    if 0<r<1000:
                                        self.data[i-1].R = r
                                    break # Exit loop once appropriate time point has been found
                    if not self.sweepFound: # Not read is also implied
                        #print(parts)
                        try:
                            (key, value) = line.split(":")
                            if key == 'Sweep 0 Start Time':
                                self.sweepStartTime = float(value.strip())
                            if key == 'Sweep 0 End Time':
                                self.sweepEndTime = float(value.strip())
                            if key == 'Sweep 0 Start Value':
                                sweepStart = float(value.strip())
                            if key == 'Sweep 0 End Value':
                                sweepEnd = float(value.strip())
                                if sweepEnd != sweepStart:
                                    self.sweepFound = True
                                    print('Sweep found at ' + str(self.sweepStartTime) + ' with end time ' + str(self.sweepEndTime))
                        except ValueError:
                            continue
        except IOError:
                print('Unable to read ECR file.')
        return
    
    def clean(self):
        i = 0
        while i < len(self.data[i].R):
            if self.data[i].R == '': # If there is no resistance data, delete the entire data point
                self.data.pop(i)
            else:
                i += 1
        return
        
    def fitSweep(self):
        I = []
        V = []
        for item in self.sweepData:
            I.append(item[0])
            V.append(item[1])
        I = np.vstack([I, np.ones(len(I))]).T
        self.sweepFit = linalg.lstsq(I, V)[0]
        intercept = float(self.sweepFit[1])
        # Recaculate data based on sweep fit
        for i in range(len(self.data)):
            #print(self.data[i].V)
            if(self.data[i].V != ''): # Skip empty rows
                self.data[i].V = self.data[i].V-intercept
                try:
                    r = self.data[i].V/self.data[i].I
                except ZeroDivisionError: # I is zero
                    continue
                if 0<r<1000:
                    self.data[i].R = r
        return
                     
    def findThresholdStrain(self, resistance):
        if not(self.particleSize > 0):
            print('Threshold strain cannot be found; strain not defined for this data set.')
            return
        for idx, item in enumerate(self.data):
            if not(item.R == '') and item.R < resistance:
                self.statistics['Strain under ' + str(resistance) +' Ohm threshold'] = self.data[idx].strain
                break
        return
               
    def findResistanceAtStrain(self, s):
        #print(self.data)
        found = False
        if not(self.particleSize > 0):
            print('Strain not defined for this data set.')
            return
        for idx, datapoint in enumerate(self.data):
            #print(datapoint.strain)
            if float(datapoint.strain) > s:
                found = True
                idx1 = idx-1
                idx2 = idx+1           
                # Find closest resistance points
                try:
                    while self.data[idx1].R == '':
                        idx1 = idx1-1
                    while self.data[idx2].R == '':
                        idx2 = idx2+1
                except IndexError: # We have reached the end of the dataset:
                    print('No resistance found for ' + str(s) + ' strain, end of dataset has been reached.')
                    return
                if abs(self.data[idx1].strain-s) > 0.1 or abs(self.data[idx2].strain-s) > 0.1:
                    print('Resistance cannot be found; no resistance points close to '+ str(s) +' strain.' )
                    return
                p1 = [self.data[idx1].strain, self.data[idx1].R]
                p2 = [self.data[idx2].strain, self.data[idx2].R]
                self.statistics['Resistance at ' + str(s) + ' strain'] = self.extrapolateR(p1,p2,s)
                break
        if found == False:
            print('Resistance cannot be found, strain does not reach ' + str(s))
        return
    
    def extrapolateR(self, p1, p2, S): #p1, p2 are two [strain, resistance coodinates], S is strain you want to find R for
        a = (p2[1]-p1[1])/(p2[0]-p1[0])
        b = p1[1]-a*p1[0]
        R = a*S+b
        return R
    
    def setMaxI(self):
        self.maxI = 0
        for idx, datapoint in enumerate(self.data):
            if datapoint.f > self.data[self.maxI].f:
                self.maxI = idx
        return
    
    def setMinR(self):
        self.minR = 1000
        for idx, datapoint in enumerate(self.data):
            if datapoint.R < self.minR:
                self.statistics['Min R'] = datapoint.R
        return
    
    def setRecoveryRatio(self):
        maxd = self.data[self.maxI].d
        lastd = self.data[-1].d
        print(maxd, lastd)
        self.statistics['Recovery ratio'] = (maxd-lastd)/maxd
        return
    
    def getContact(self):
        return self.contact