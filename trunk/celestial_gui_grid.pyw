#! /usr/bin/python3
# -*- coding: utf-8 -*-
about = """
Written for Python 3.4+ (any 3+ will do as long as you do not need pip on windows)
Uses Ephem and ReportLab (optional, for PDF output)
This script works without reportlab, but you will not be able to generate the tables as PDF
Get packages with pip install reportlab and pip install ephem

DO NOT USE FOR ACTUAL NAVIGATION!

The MIT License (MIT)

Copyright (c) 2014 Lars Neumann

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""

import ephem
import sys, os, time
import pickle
from math import *
import tkinter as tk
import tkinter.messagebox as messagebox
import tkinter.filedialog as filedialog
from tkinter.ttk import Progressbar
from datetime import *
from collections import OrderedDict
import operator
# for star sorting + selection
import itertools
try:
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.platypus import Table
    from reportlab.lib import colors
    from reportlab.lib.enums import *

    from reportlab.platypus import BaseDocTemplate, SimpleDocTemplate, Paragraph, Spacer, Frame, PageTemplate, NextPageTemplate, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.rl_config import defaultPageSize
    from reportlab.lib.units import inch
except ImportError as e:
    reportlab_error = e
else:
    reportlab_error = None

"""
List of stars used for Celestial Navigation
"""
db = """\
Achernar,f|V|B3,1:37:42.9,-57:14:12,0.46,2000
Acrux,f|M|B0,12:26:35.9,-63:5:57,1.33,2000
Aldebaran,f|M|K5,4:35:55.2,16:30:33,0.85,2000
Altair,f|M|A7,19:50:47,8:52:6,0.77,2000
Antares,f|D|M1,16:29:24.4,-26:25:55,0.96,2000
Arcturus,f|V|K1,14:15:39.7,19:10:57,-0.04,2000
Betelgeuse,f|M|M1,5:55:10.3,7:24:25,0.5,2000
Canopus,f|S|F0,6:23:57.1,-52:41:45,-0.72,2000
Capella,f|M|G5,5:16:41.4,45:59:53,0.08,2000
Deneb,f|D|A2,20:41:25.9,45:16:49,1.25,2000
Fomalhaut,f|V|A3,22:57:39.1,-29:37:20,1.16,2000
Pollux,f|M|K0,7:45:18.9,28:1:34,1.14,2000
Procyon,f|M|F5,7:39:18.1,5:13:30,0.38,2000
Regulus,f|M|B7,10:8:22.3,11:58:2,1.35,2000
Rigel,f|M|B8,5:14:32.3,-8:12:6,0.12,2000
Rigil Kent.,f|M|G2,14:39:35.9,-60:50:7,-0.01,2000
Sirius,f|M|A1,6:45:8.9,-16:42:58,-1.46,2000
Spica,f|M|B1,13:25:11.6,-11:9:41,0.98,2000
Vega,f|M|A0,18:36:56.3,38:47:1,0.03,2000
Acamar,f|D|A4,2:58:15.7,-40:18:17,3.24,2000
Alioth,f|V|A0,12:54:1.7,55:57:35,1.77,2000
Alkaid,f|V|B3,13:47:32.4,49:18:48,1.86,2000
Alphard,f|M|K3,9:27:35.2,-8:39:31,1.98,2000
Alphecca,f|V|A0,15:34:41.3,26:42:53,2.23,2000
Alpheratz,f|D|B8,0:8:23.3,29:5:26,2.06,2000
Denebola,f|M|A3,11:49:3.6,14:34:19,2.14,2000
Diphda,f|V|G9,0:43:35.4,-17:59:12,2.04,2000
Dubhe,f|D|K0,11:3:43.7,61:45:3,1.79,2000
Enif,f|M|K2,21:44:11.2,9:52:30,2.39,2000
Gienah,f|V|B8,12:15:48.4,-17:32:31,2.59,2000
Hamal,f|V|K2,2:7:10.4,23:27:45,2,2000
Kochab,f|D|K4,14:50:42.3,74:9:20,2.08,2000
Menkar,f|V|M1,3:2:16.8,4:5:23,2.53,2000
Miaplacidus,f|S|A2,9:13:12,-69:43:2,1.68,2000
Mirfak,f|D|F5,3:24:19.4,49:51:40,1.79,2000
Nunki,f|D|B2,18:55:15.9,-26:17:48,2.02,2000
Peacock,f|M|B2,20:25:38.9,-56:44:6,1.94,2000
Rasalhague,f|B|A5,17:34:56.1,12:33:36,2.08,2000
Schedar,f|M|K0,0:40:30.5,56:32:14,2.23,2000
Shaula,f|M|B2,17:33:36.5,-37:6:14,1.63,2000
Suhail,f|D|K4,9:7:59.8,-43:25:57,2.21,2000
Polaris,f|M|F7,2:31:48.7,89:15:51,2.02,2000
"""
# These stars are not used in the original Pub 249
# Use them as you like
"""
Hadar,f|D|B1,14:3:49.4,-60:22:23,0.61,2000
Cep Gamma,f|V|K1,23:39:20.8,77:37:57,3.21,2000
Markab,f|V|B9,23:4:45.7,15:12:19,2.49,2000
Scheat,f|M|M2,23:3:46.5,28:4:58,2.42,2000
Gru Beta,f|V|M5,22:42:40.1,-46:53:5,2.1,2000
Tuc Alpha,f|B|K3,22:18:30.1,-60:15:35,2.86,2000
Al Na'ir,f|D|B7,22:8:14,-46:57:40,1.74,2000
Cap Delta,f|M|Am,21:47:2.4,-16:7:38,2.87,2000
Aqu Beta,f|M|G0,21:31:33.5,-5:34:16,2.91,2000
Alderamin,f|M|A7,21:18:34.8,62:35:8,2.44,2000
Cyg Epsilon,f|M|K0,20:46:12.7,33:58:13,2.46,2000
Ind Alpha,f|M|K0,20:37:34,-47:17:29,3.11,2000
Cyg Gamma,f|M|F8,20:22:13.7,40:15:24,2.2,2000
Aql Gamma,f|D|K3,19:46:15.6,10:36:48,2.72,2000
Cyg Delta,f|M|B9,19:44:58.5,45:7:51,2.87,2000
Albireo,f|M|K3,19:30:43.3,27:57:35,3.08,2000
Sgr Pi,f|M|F2,19:9:45.8,-21:1:25,2.89,2000
Aql Zeta,f|M|A0,19:5:24.6,13:51:48,2.99,2000
Sgr Zeta,f|M|A2,19:2:36.7,-29:52:49,2.6,2000
Sgr Lambda,f|S|K1,18:27:58.2,-25:25:18,2.81,2000
Kaus Australis,f|D|B9,18:24:10.3,-34:23:5,1.85,2000
Sgr Delta,f|M|K3,18:20:59.7,-29:49:41,2.7,2000
Sgr Gamma,f|V|K0,18:5:48.5,-30:25:27,2.99,2000
Eltanin,f|M|K5,17:56:36.4,51:29:20,2.23,2000
Oph Beta,f|S|K2,17:43:28.4,4:34:2,2.77,2000
Sco Kappa,f|V|B1,17:42:29.3,-39:1:48,2.41,2000
Sco Theta,f|V|F1,17:37:19.2,-42:59:52,1.87,2000
Ara Alpha,f|D|B2,17:31:50.5,-49:52:34,2.95,2000
Sco Upsilo,f|S|B2,17:30:45.8,-37:17:45,2.69,2000
Dra Beta,f|M|G2,17:30:26,52:18:5,2.79,2000
Ara Beta,f|S|K3,17:25:18,-55:31:48,2.85,2000
Her Alpha,f|M|M5,17:14:38.9,14:23:25,3.48,2000
Sabik,f|M|A2,17:10:22.7,-15:43:29,2.43,2000
Ara Zeta,f|S|K3,16:58:37.2,-55:59:25,3.13,2000
Sco Epsilo,f|V|K2,16:50:9.8,-34:17:36,2.29,2000
Atria,f|S|K2,16:48:39.9,-69:1:40,1.92,2000
Her Zeta,f|D|G0,16:41:17.2,31:36:11,2.81,2000
Oph Zeta,f|V|O9,16:37:9.5,-10:34:2,2.56,2000
Sco Tau,f|S|B0,16:35:53,-28:12:58,2.82,2000
Her Beta,f|D|G7,16:30:13.2,21:29:23,2.77,2000
Antares,f|D|M1,16:29:24.4,-26:25:55,0.96,2000
Dra Eta,f|M|G8,16:23:59.5,61:30:51,2.74,2000
Oph Delta,f|D|M0,16:14:20.7,-3:41:40,2.74,2000
Sco Beta1,f|M|B1,16:5:26.2,-19:48:20,2.62,2000
Dschubba,f|M|B0,16:0:20,-22:37:18,2.32,2000
Sco Pi,f|M|B1,15:58:51.1,-26:6:51,2.89,2000
TrA Beta,f|D|F2,15:55:8.5,-63:25:50,2.85,2000
Ser Alpha,f|M|K2,15:44:16.1,6:25:32,2.65,2000
Lup Gamma,f|D|B2,15:35:8.5,-41:10:1,2.78,2000
UMi Gamma,f|V|A3,15:20:43.7,71:50:2,3.05,2000
TrA Gamma,f|V|A1,15:18:54.6,-68:40:46,2.89,2000
Lib Beta,f|V|B8,15:17:0.4,-9:22:59,2.61,2000
Lup Beta,f|S|B2,14:58:31.9,-43:8:2,2.68,2000
Zubenelgenubi,f|M|A3,14:50:52.7,-16:2:30,2.75,2000
Boo Epsilon,f|M|K0,14:44:59.2,27:4:27,2.7,2000
Lup Alpha,f|D|B1,14:41:55.8,-47:23:18,2.3,2000
Cen Eta,f|V|B1,14:35:30.4,-42:9:28,2.31,2000
Boo Gamma,f|M|A7,14:32:4.7,38:18:30,3.03,2000
Arcturus,f|V|K1,14:15:39.7,19:10:57,-0.04,2000
Menkent,f|D|K0,14:6:41,-36:22:12,2.06,2000
Cen Zeta,f|S|B2,13:55:32.4,-47:17:18,2.55,2000
Boo Eta,f|D|G0,13:54:41.1,18:23:52,2.68,2000
Cen Epsilon,f|D|B1,13:39:53.2,-53:27:59,2.3,2000
Mizar,f|M|A1,13:23:55.5,54:55:31,2.27,2000
Cen Iota,f|S|A2,13:20:35.8,-36:42:44,2.75,2000
Vir Epsilon,f|D|G8,13:2:10.6,10:57:33,2.83,2000
Cor Caroli,f|D|A0,12:56:1.7,38:19:6,2.9,2000
Mimosa,f|M|B0,12:47:43.2,-59:41:19,1.25,2000
Vir Gamma,f|M|F0,12:41:39.6,-1:26:58,3.65,2000
Muhlifain,f|M|A1,12:41:31,-48:57:35,2.17,2000
Mus Alpha,f|D|B2,12:37:11,-69:8:8,2.69,2000
Crv Beta,f|V|G5,12:34:23.2,-23:23:48,2.65,2000
Gacrux,f|M|M3,12:31:9.9,-57:6:48,1.63,2000
Cen Delta,f|M|B2,12:8:21.5,-50:43:21,2.6,2000
Phecda,f|V|A0,11:53:49.8,53:41:41,2.44,2000
Leo Delta,f|M|A4,11:14:6.5,20:31:25,2.56,2000
UMa Psi,f|S|K1,11:9:39.8,44:29:55,3.01,2000
Merak,f|V|A1,11:1:50.5,56:22:57,2.37,2000
Vel Mu,f|D|G5,10:46:46.2,-49:25:12,2.69,2000
Car Theta,f|S|B0,10:42:57.4,-64:23:40,2.76,2000
Algeiba,f|M|K1,10:19:58.3,19:50:30,2.61,2000
Leo Epsilo,f|V|G1,9:45:51.1,23:46:27,2.98,2000
Vel Nu,f|V|K5,9:31:13.3,-57:2:4,3.13,2000
Vel Kappa,f|S|B2,9:22:6.8,-55:0:39,2.5,2000
Car Iota,f|V|A8,9:17:5.4,-59:16:31,2.25,2000
UMa Iota,f|M|A7,8:59:12.4,48:2:30,3.14,2000
Vel Delta,f|M|A1,8:44:42.2,-54:42:30,1.96,2000
Avior,f|V|K3,8:22:30.8,-59:30:35,1.86,2000
Vel Gamma,f|M|WC,8:9:32,-47:20:12,1.78,2000
Pup Rho,f|D|F6,8:7:32.6,-24:18:15,2.81,2000
Pup Zeta,f|S|O5,8:3:35.1,-40:0:12,2.25,2000
Castor,f|M|A1,7:34:36,31:53:18,1.98,2000
Pup Sigma,f|D|K5,7:29:13.8,-43:18:5,3.25,2000
CMi Beta,f|M|B8,7:27:9,8:17:22,2.9,2000
CMa Eta,f|D|B5,7:24:5.7,-29:18:11,2.45,2000
Pup Pi,f|D|K3,7:17:8.6,-37:5:51,2.7,2000
Wezen,f|V|F8,7:8:23.5,-26:23:36,1.84,2000
CMa Omicr,f|S|B3,7:3:1.5,-23:50:0,3.02,2000
Adhara,f|D|B2,6:58:37.5,-28:58:20,1.5,2000
Pup Tau,f|S|K1,6:49:56.2,-50:36:53,2.93,2000
Alhena,f|M|A0,6:37:42.7,16:23:57,1.93,2000
Mirzam,f|D|B1,6:22:42,-17:57:21,1.98,2000
Aur Theta,f|M|A0,5:59:43.3,37:12:45,2.62,2000
Menkalinan,f|M|A2,5:59:31.7,44:56:51,1.9,2000
Ori Kappa,f|V|B0,5:47:45.4,-9:40:11,2.06,2000
Alnitak,f|M|O9,5:40:45.5,-1:56:34,2.05,2000
Phact,f|D|B7,5:39:38.9,-34:4:27,2.64,2000
Tau Zeta,f|B|B4,5:37:38.7,21:8:33,3,2000
Alnilam,f|D|B0,5:36:12.8,-1:12:7,1.7,2000
Ori Iota,f|M|O9,5:35:26,-5:54:36,2.77,2000
Lep Alpha,f|M|F0,5:32:43.8,-17:49:20,2.58,2000
Ori Delta,f|M|O9,5:32:0.4,-0:17:57,2.23,2000
Lep Beta,f|M|G5,5:28:14.7,-20:45:34,2.84,2000
Elnath,f|D|B7,5:26:17.5,28:36:27,1.65,2000
Bellatrix,f|D|B2,5:25:7.9,6:20:59,1.64,2000
Eri Beta,f|D|A3,5:7:51,-5:5:11,2.79,2000
Aur Iota,f|V|K3,4:56:59.6,33:9:58,2.69,2000
Eri Gamma,f|D|M0,3:58:1.8,-13:30:31,2.95,2000
Per Epsilo,f|M|B0,3:57:51.2,40:0:37,2.89,2000
Per Zeta,f|M|B1,3:54:7.9,31:53:1,2.85,2000
Alcyone,f|M|B7,3:47:29.1,24:6:18,2.87,2000
Algol,f|M|B8,3:8:10.1,40:57:20,2.12,2000
Tri Beta,f|S|A5,2:9:32.6,34:59:14,3,2000
Almak,f|M|K3,2:3:54,42:19:47,2.26,2000
Hyi Alpha,f|S|F0,1:58:46.2,-61:34:11,2.86,2000
Sheratan,f|V|A5,1:54:38.4,20:48:29,2.64,2000
Ruchbah,f|D|A5,1:25:49,60:14:7,2.68,2000
Mirach,f|M|M0,1:9:43.9,35:37:14,2.06,2000
Cas Gamma,f|M|B0,0:56:42.5,60:43:0,2.47,2000
Ankaa,f|B|K0,0:26:17,-42:18:22,2.39,2000
Hyi Beta,f|V|G2,0:25:45.1,-77:15:15,2.8,2000
Algenib,f|M|B2,0:13:14.2,15:11:1,2.83,2000
Caph,f|D|F2,0:9:10.7,59:8:59,2.27,2000
"""

class DEBUG:
    __state = False

    @staticmethod
    def ON():
        return DEBUG.__state

    @staticmethod
    def set(state):
        if state==1:
            DEBUG.__state=True
        else:
            DEBUG.__state=False

def roundto(x, base=4):
    return int(base * round(float(x)/base))

def dist(lat1, lon1, lat2, lon2):
    return 2*asin(sqrt(pow(sin((lat1-lat2)/2),2) + cos(lat1)*cos(lat2)*pow(sin((lon1-lon2)/2),2)))

def bear_dist(lat, lon, brg, distance):
    newlat=asin(sin(lat)*cos(distance)+cos(lat)*sin(distance)*cos(brg))
    dlon=atan2(sin(brg)*sin(distance)*cos(lat),cos(distance)-sin(lat)*sin(lat))
    newlon=ephem.degrees(lon+dlon).znorm
    return(ephem.degrees(newlat), ephem.degrees(newlon))

def find_int(lat1, lon1, crs13, lat2, lon2, crs23):
    dst12=2*asin(sqrt(
        (pow(sin((lat1-lat2)/2),2)+
                       cos(lat1)*cos(lat2)*pow(sin((-lon1+lon2)/2),2))))
    if(sin(-lon2+lon1)<0):
       crs12=acos((sin(lat2)-sin(lat1)*cos(dst12))/(sin(dst12)*cos(lat1)))
       crs21=2.*pi-acos((sin(lat1)-sin(lat2)*cos(dst12))/(sin(dst12)*cos(lat2)))
    else:
       crs12=2.*pi-acos((sin(lat2)-sin(lat1)*cos(dst12))/(sin(dst12)*cos(lat1)))
       crs21=acos((sin(lat1)-sin(lat2)*cos(dst12))/(sin(dst12)*cos(lat2)))

    ang1=ephem.degrees(crs13-crs12).znorm
    ang2=ephem.degrees(crs21-crs23).znorm

    if(sin(ang1)==0 and sin(ang2)==0):
        raise Exception("infinity of intersections")
    else:
        if(sin(ang1)*sin(ang2)<0):
            raise Exception("intersection ambiguous")
        else:
           ang1=abs(ang1)
           ang2=abs(ang2)
           ang3=acos(-cos(ang1)*cos(ang2)+sin(ang1)*sin(ang2)*cos(dst12)) 
           dst13=atan2(sin(dst12)*sin(ang1)*sin(ang2),cos(ang2)+cos(ang1)*cos(ang3))
           lat3=asin(sin(lat1)*cos(dst13)+cos(lat1)*sin(dst13)*cos(crs13))
           dlon=atan2(sin(crs13)*sin(dst13)*cos(lat1),cos(dst13)-sin(lat1)*sin(lat3))
           lon3=ephem.degrees(lon1+dlon).znorm
    return(ephem.degrees(lat3), ephem.degrees(lon3))


def find_int3(star1, star2, star3):
    # mit exception handling richtigen Schnittpunkt finden
    options = ((+0.5*pi, +0.5*pi),(+0.5*pi, -0.5*pi),(-0.5*pi, -0.5*pi),(-0.5*pi, +0.5*pi))
    combinations = ((star1,star2),(star1,star3),(star2, star3))
    intersections = []
    # 3 Schnittpunkte bestimmen (1-2, 1-3, 2-3)
    for shot1, shot2 in combinations:
        lat = None
        lon = None
        i=0
        while lat is None:
            try:
                lat,lon = find_int(shot1.int_lat, shot1.int_lon, ephem.degrees(shot1.zn+options[i][0]).norm,
                                   shot2.int_lat, shot2.int_lon, ephem.degrees(shot2.zn+options[i][1]).norm)
                if dist(lat, lon, shot1.int_lat, shot1.int_lon)>(pi/(180*60))*200: # intersection more than 200 miles away?
                    lat = None
                    raise Exception('too far')
            except Exception as e:
                if i>3:
                    raise Exception('no valid intersection found')
                i+=1
            else:
                if DEBUG.ON():
                    print('Intersection',shot1.name,'+',shot2.name,':',lat,lon)
                intersections.append((lat,lon))
    return intersections
    
def least_error(intersections):
    # valid only for flat triangle...
    if len(intersections)!=3:
        raise Exception('not a triangle')
    lat = 0
    lon = 0
    for ints in intersections:
        lat += ints[0]
        lon += ints[1]
    lat /= len(intersections)
    lon /= len(intersections)
    return (ephem.degrees(lat),ephem.degrees(lon))

def s2d(v, t=60):
    """
    Convert a speed v and a time t in minutes to a distance in rad
    If time is omitted it defaults to 1 hour
    """
    return pi*v*t/648000

class MyDocTemplate(BaseDocTemplate):
    """Override the BaseDocTemplate class to do custom handle_XXX actions"""

    def __init__(self, *args, **kwargs):
        BaseDocTemplate.__init__(self, *args, **kwargs)
        self.latitude = None
        self.PAGE_WIDTH, self.PAGE_HEIGHT = A4

    def afterPage(self):
        """Called after each page has been processed"""
        if self.latitude is None:
            return
        # saveState keeps a snapshot of the canvas state, so you don't
        # mess up any rendering that platypus will do later.
        self.canv.saveState()

        # Reset the origin to (0, 0), remember, we can restore the
        # state of the canvas later, so platypus should be unaffected.
        self.canv._x = 0
        self.canv._y = 0
        if self.latitude < 0:
            hemisphere = 'S'
        else:
            hemisphere = 'N'
        self.canv.setFont('Helvetica-Bold',12)
        self.canv.drawString(inch/3, self.PAGE_HEIGHT-0.4 * inch, "LAT %d°%s" % (abs(self.latitude), hemisphere))
        self.canv.drawRightString(self.PAGE_WIDTH-inch/3, self.PAGE_HEIGHT-0.4 * inch, "LAT %d°%s" % (abs(self.latitude), hemisphere))    

        # Now we restore the canvas back to the way it was.
        self.canv.restoreState()

    def afterFlowable(self, flowable):
        if hasattr(flowable,'latitude'):
            self.latitude = flowable.latitude

class TablesPDF:
    def __init__(self):
        if reportlab_error:
            messagebox.showinfo("Error", "ReportLab could not be imported!")
            raise reportlab_error

        self.format = A4
        self.PAGE_WIDTH, self.PAGE_HEIGHT = self.format

        self.style = getSampleStyleSheet()["Normal"]
        self.styleCentered = getSampleStyleSheet()["Normal"]
        self.styleCentered.alignment=TA_CENTER
        self.styleTable = getSampleStyleSheet()["Normal"]
        self.styleTable.alignment=TA_CENTER
        self.styleTable.fontName='Helvetica'
        self.styleTable.fontSize=6

        self.Title = 'PUB. NO. 249 VOL. 1'
        self.Author = 'Lars Neumann'
        self.pageinfo = 'SIGHT REDUCTION TABLES FOR AIR NAVIGATION'
        self.epoch = 2010

    def title(self, canvas, doc):
        canvas.saveState()
        canvas.setFont('Times-Bold',16)
        canvas.drawCentredString(self.PAGE_WIDTH/2.0, doc.height-108, self.Title)
        canvas.drawCentredString(self.PAGE_WIDTH/2.0, doc.height-128, self.pageinfo)
        canvas.setFont('Times-Roman',12)
        canvas.drawCentredString(self.PAGE_WIDTH/2.0, doc.height-148, '(SELECTED STARS)')
        canvas.drawCentredString(self.PAGE_WIDTH/2.0, doc.height-168, 'Epoch '+str(self.epoch)+'.0')    
        canvas.drawCentredString(self.PAGE_WIDTH/2.0, doc.height-188, 'Author '+self.Author)
        canvas.drawCentredString(self.PAGE_WIDTH/2.0, doc.height-250, 'Do not use for navigation!')
        canvas.drawCentredString(self.PAGE_WIDTH/2.0, doc.height-270, 'Use at your own risk - although all data is believed to be correct use approved documents for actual navigation.')
        canvas.drawCentredString(self.PAGE_WIDTH/2.0, doc.height-290, 'You have been warned!')
        canvas.restoreState()

    def normalPagesFooter(self, canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica',9)
        canvas.drawCentredString(self.PAGE_WIDTH/2.0, 0.3 * inch, "%d" % (doc.page))
        canvas.restoreState()

    def copyObserver(self, observer):
        temp_observer = ephem.Observer()
        temp_observer.lat = observer.lat
        temp_observer.lon = observer.lon
        temp_observer.pressure = observer.pressure
        temp_observer.date = observer.date
        temp_observer.epoch = observer.epoch
        return temp_observer

    def __new_lha_zero_observer(self):
        temp_observer = ephem.Observer()
        temp_observer.lat = ephem.degrees(0)
        temp_observer.lon = ephem.degrees(0)
        temp_observer.pressure = 0
        temp_observer.date = ephem.date(self.epoch)
        temp_observer.epoch = ephem.date(self.epoch)
        # get the lon where lha_a is 0/360
        temp_observer.lon = -ephem.degrees(temp_observer.sidereal_time())
        return temp_observer

    def __get_str_degrees(self, angle, decimals = 0, digits = 3, signs = True):
            angle_ddd = int(angle*180/pi)
            mm = ((angle*180/pi)-angle_ddd)*60
            if decimals > 0:
                angle_mm = round(mm,decimals)
            else:
                angle_mm = round(mm)
            if signs:
                return str(angle_ddd).zfill(digits)+'° '+str(angle_mm).zfill(2)+'´'
            else:
                return str(angle_ddd).zfill(digits)+' '+str(angle_mm).zfill(2)

    def __computeLat(self, lat):
        """
        returns the table for the given latitude, a set of stars is given for
        each 15/30 deg of LHA
        """
        data = []
        header = ['LHA\nAries', 'Hc', 'Zn', 'Hc', 'Zn', 'Hc', 'Zn', 'Hc', 'Zn', 'Hc', 'Zn', 'Hc', 'Zn', 'Hc', 'Zn']
        data.append(header)
        # for lat > 69 determine a new set of stars for each 30 deg, otherwise 15 deg
        if abs(lat) > 69:
            base_range = 12
            delta_range = 30
            mid_point = ephem.hours('1:00')
            length = ephem.hours('2:00')
            # used to fill up the frame
            add_lines = 3
        else:
            base_range = 24
            delta_range = 15
            mid_point = ephem.hours('0:30')
            length = ephem.hours('1:00')
            add_lines = 0
        for lha_base in range(base_range):
            observer = self.__new_lha_zero_observer()
            observer.lon += lha_base * length
            # get a set of 7 bodies with favorable Hc/Zn/mag
            mid_point_obs = self.copyObserver(observer)
            # go to mid point of the 15 deg block
            mid_point_obs.lon += mid_point
            star = []
            number = 7
            try:
                stars, threestarfix = GoodSet(mid_point_obs).compute(number = number)
            except StopIteration as e:
                print('LHA A:',ephem.degrees(mid_point_obs.sidereal_time()),e)
            else:
                stars.sort(key=operator.attrgetter('az'))
                block_header = ['']
                for body in stars:
                    # uppercase name for magnitude brighter 1.5
                    if body.mag < 1.5:
                        name = body.name.upper()
                    else:
                        name = body.name
                    # add * to bodies for 3 star fix
                    if body.name in threestarfix:
                        block_header.append('\u25c6'+name)
                    else:
                        block_header.append(name)
                    # span 2 columns
                    block_header.append('')
                data.append(block_header)
                for lha_delta in range(delta_range):
                    data_row = []
                    temp_observer = self.copyObserver(observer)
                    temp_observer.lon += lha_delta * ephem.degrees('1')
                    lha_a = round(ephem.degrees(temp_observer.sidereal_time())*180/pi)
                    if lha_a == 360:
                        lha_a = 0
                    data_row.append(lha_a)
                    for body in stars:
                        body.compute(temp_observer)
                        alt_dd = int(body.alt*180/pi)
                        alt_mm = round(((body.alt*180/pi)-alt_dd)*60)
                        az = round(body.az*180/pi)
                        data_row.append(str(alt_dd)+' '+str(alt_mm).zfill(2))
                        data_row.append(str(az).zfill(3))
                    data.append(data_row)
                # pad rows to fill up frame
                if lha_base % (base_range/4) == 2:
                    for a in range(add_lines):
                        data.append('')
        style =[('FONT',(0,0),(-1,-1),'Helvetica-Bold', 6),
                ('FONT',(1,1),(-1,-1),'Helvetica', 5),
                #('LINEBELOW',(0,0),(-1,0),0.2,colors.black),
                ('BOX',(0,0),(0,-1),0.2,colors.black),
                ('LEFTPADDING',(0,0),(-1,-1),0),
                ('RIGHTPADDING',(0,0),(-1,-1),0),
                ('BOTTOMPADDING',(0,0),(-1,-1),0),
                ('TOPPADDING',(0,0),(-1,-1),0),
                ('ALIGN', (0,0), (-1,-1), 'CENTRE'),
                ('VALIGN',(0,0), (-1,-1), 'TOP'),
                ('SPAN',(0,0),(0,1))]
                #('LINEBELOW',(0,1),(0,1),0.2,colors.black)]
        for column in range(1,8):
            style.append(('BOX',(column*2-1,0),(column*2,-1),0.2,colors.black))
            padded_rows = 0
            for row in range(base_range):
                if row % (base_range/4) == 0:
                    if abs(lat) > 69:
                        # we have to account for the padded rows to fill up the frame > +-69 lat
                        padded_rows=int(row*12/base_range)
                    style.append(('LINEBELOW',(0,(delta_range+1)*row+padded_rows),(-1,(delta_range+1)*row+padded_rows),0.2,colors.black))
                    style.append(('LINEBELOW',(0,(delta_range+1)*row+padded_rows+1),(0,(delta_range+1)*row+padded_rows+1),0.2,colors.black))
                # we have to account for the padded rows to fill up the frame > +-69 lat
                style.append(('SPAN',(column*2-1,(delta_range+1)*row+padded_rows+1),(column*2,(delta_range+1)*row+padded_rows+1)))
                style.append(('FONT',(0,(delta_range+1)*row+padded_rows+1),(-1,(delta_range+1)*row+padded_rows+1),'Helvetica-Bold', 5))
        t=Table(data, repeatRows = 1, style = style,
                colWidths = self.PAGE_WIDTH/(2.0*15)-2,
                rowHeights = 8)
        return(t)

    def __compute_lat_table(self, progress_bar, lat_min, lat_max):
        Story = []
        Story.append(NextPageTemplate('Table'))
        
        for lat in range(lat_max,lat_min-1,-1):
            Story.append(PageBreak())
            if progress_bar is not None:
                progress_bar['value'] = ((lat_max-lat)/(lat_max+1-lat_min)*100)
                progress_bar.update_idletasks()
            table = self.__computeLat(lat)
            # empty paragraph just to assign the latitude for the header
            p = Paragraph('', self.style)
            Story.append(p)
            Story[-1].latitude = lat
            Story.append(table)
        return Story

    def __table_gha_year(self):
        gha_observer = ephem.Observer()
        gha_observer.lat = ephem.degrees(0)
        gha_observer.lon = ephem.degrees(0)
        gha_observer.pressure = 0
        gha_observer.date = ephem.date(self.epoch)
        gha_observer.epoch = ephem.date(self.epoch)

        data = []
        header = ['Year', 'Jan. 1', 'Feb. 1', 'Mar. 1', 'Apr. 1', 'May 1', 'June 1', 'July 1', 'Aug. 1', 'Sept. 1', 'Oct. 1', 'Nov. 1', 'Dec. 1']
        data.append(header)
        header = ['', '°   ´', '°   ´', '° ´', '° ´', '° ´', '° ´', '° ´', '° ´', '° ´', '° ´', '° ´', '° ´']
        data.append(header)
        for year in range(self.epoch.tuple()[0]-4, self.epoch.tuple()[0]+5):
            if year == self.epoch:
                data.append('')
            data_row = [year]
            for month in range(1,13):
                gha_observer.date = ephem.date((year,month))
                gha = ephem.degrees(gha_observer.sidereal_time())
                # gather data for the example
                if month == 8 and year == self.epoch.tuple()[0]+2:
                    example = gha
                gha_ddd = int(gha*180/pi)
                gha_mm = round(((gha*180/pi)-gha_ddd)*60)
                data_row.append(str(gha_ddd).zfill(3)+'  '+str(gha_mm).zfill(2))
            data.append(data_row)
        style =[('FONT',(0,0),(-1,-1),'Helvetica', 6),
                ('FONT',(0,1),(0,-1),'Helvetica-Bold', 6),
                ('LINEBELOW',(0,0),(-1,0),0.2,colors.black),
                ('LEFTPADDING',(0,0),(-1,-1),0),
                ('RIGHTPADDING',(0,0),(-1,-1),0),
                ('BOTTOMPADDING',(0,0),(-1,-1),0),
                ('TOPPADDING',(0,0),(-1,-1),0),
                ('ALIGN', (0,0), (-1,-1), 'CENTRE'),
                ('VALIGN',(0,0), (-1,-1), 'MIDDLE'),
                ('VALIGN',(0,1), (-1,1), 'BOTTOM')]
        for column in range(13):
            style.append(('BOX',(column-1,0),(column,-1),0.2,colors.black))
        style.append(('BOX',(0,0),(-1,-1),2,colors.black))
        t=Table(data, style = style,
                colWidths = self.PAGE_WIDTH/13*2/3,
                rowHeights = 10)
        return t, example

    def __table_gha_days(self):
        data = []
        # 365.2422 days for 360 deg
        sidereal_day = ephem.degrees(2*pi/365.2422)
        # 23.9344699 hours for 360 deg
        sidereal_hour = ephem.degrees(2*pi/23.9344699)
        # table has 2 blocks, first from day 1-16, second from 17-32
        for block in range(2):
            header = ['Day']
            header2 = ['h']
            for day in range(16):
                header.append(day+1+block*16)
                header2.append('°  ´')
            data.append(header)
            data.append(header2)
            for hour in range(24):
                data_row = [str(hour).zfill(2)]
                for day in range(16):
                    gha = ephem.degrees((sidereal_day*(day+block*16))+(sidereal_hour*hour)).norm
                    # gather data for the example
                    if hour == 5 and (day+block*16) == 16:
                        example = gha
                    gha_ddd = int(gha*180/pi)
                    gha_mm = round(((gha*180/pi)-gha_ddd)*60)
                    data_row.append(str(gha_ddd)+'  '+str(gha_mm).zfill(2))
                data.append(data_row)
        style =[('FONT',(0,0),(-1,-1),'Helvetica', 6),
                ('FONT',(1,0),(-1,0),'Helvetica-Bold', 6),
                ('FONT',(1,26),(-1,26),'Helvetica-Bold', 6),
                ('FONT',(0,2),(0,25),'Helvetica-Bold', 6),
                ('FONT',(0,28),(0,-1),'Helvetica-Bold', 6),
                ('LINEBELOW',(0,0),(-1,0),0.2,colors.black),
                ('BOX',(0,26),(-1,26),0.2,colors.black),
                ('LEFTPADDING',(0,0),(-1,-1),4),
                ('RIGHTPADDING',(0,0),(-1,-1),4),
                ('BOTTOMPADDING',(0,0),(-1,-1),0),
                ('TOPPADDING',(0,0),(-1,-1),0),
                ('ALIGN', (0,0), (-1,-1), 'CENTRE'),
                ('ALIGN', (1,2), (-1,25), 'RIGHT'),
                ('ALIGN', (1,28), (-1,-1), 'RIGHT'),
                ('VALIGN',(0,0), (-1,-1), 'MIDDLE'),
                ('VALIGN',(0,1), (-1,1), 'BOTTOM'),
                ('VALIGN',(0,27), (-1,27), 'BOTTOM')]
        for column in range(17):
            style.append(('BOX',(column,0),(column+1,-1),0.2,colors.black))
        style.append(('BOX',(0,0),(-1,-1),2,colors.black))
        t=Table(data, style = style,
                colWidths = self.PAGE_WIDTH/17*0.8,
                rowHeights = 10)
        return t, example

    def __table_gha_minutes(self):
        data = []
        data.append('')
        # 23.9344699 hours for 360 deg / 60 minutes
        sidereal_minute = ephem.degrees(2*pi/23.9344699/60)
        header = ['']
        header2 = ['m']
        for seconds in range(0, 61, 4):
            p = Paragraph('<b>'+str(seconds).zfill(2)+'<super>s</super></b>', self.styleTable)
            header.append(p)
            header2.append('°  ´')
            if seconds == 28 or seconds ==60:
                header.append('')
                header2.append('m')
        data.append(header)
        data.append(header2)
        for minute in range(60):
            data_row = [str(minute).zfill(2)]
            for seconds in range(0, 61, 4):
                gha = ephem.degrees((sidereal_minute*minute)+(sidereal_minute/60*seconds))
                # gather data for the example
                if minute == 11 and seconds == 40:
                    example = gha
                gha_ddd = int(gha*180/pi)
                gha_mm = round(((gha*180/pi)-gha_ddd)*60)
                data_row.append(str(gha_ddd)+' '+str(gha_mm).zfill(2))
                if seconds == 28 or seconds ==60:
                    data_row.append(str(minute).zfill(2))
            data.append(data_row)
            if minute % 5 == 4:
                data.append('')
        style =[('FONT',(0,0),(-1,-1),'Helvetica', 6),
                ('FONT',(0,1),(-1,1),'Helvetica-Bold', 6),
                ('FONT',(0,3),(0,-1),'Helvetica-Bold', 6),
                ('FONT',(9,3),(9,-1),'Helvetica-Bold', 6),
                ('FONT',(18,3),(18,-1),'Helvetica-Bold', 6),
                ('BOX',(0,0),(0,-1),0.2,colors.black),
                ('BOX',(9,0),(9,-1),0.2,colors.black),
                ('BOX',(18,0),(18,-1),0.2,colors.black),
                ('LINEAFTER',(4,0),(4,-1),0.2,colors.black),
                ('LINEAFTER',(13,0),(13,-1),0.2,colors.black),
                ('LEFTPADDING',(0,0),(-1,-1),6),
                ('RIGHTPADDING',(0,0),(-1,-1),6),
                ('BOTTOMPADDING',(0,0),(-1,-1),0),
                ('TOPPADDING',(0,0),(-1,-1),0),
                ('ALIGN', (0,0), (-1,-1), 'CENTRE'),
                ('ALIGN', (1,3), (8,-1), 'RIGHT'),
                ('ALIGN', (10,3), (-2,-1), 'RIGHT'),
                ('VALIGN',(0,0), (-1,-1), 'MIDDLE'),
                ('VALIGN',(0,0), (-1,0), 'TOP'),
                ('VALIGN',(0,2), (-1,1), 'TOP')]
        for column in range(17):
            style.append(('BOX',(column+1,0),(column+1,1),0.2,colors.black))
        style.append(('BOX',(0,0),(-1,-1),2,colors.black))
        t=Table(data, style = style,
                colWidths = self.PAGE_WIDTH/17*0.8,
                rowHeights = 8)
        return t, example

    def __compute_gha_table(self):
        Story = []
        Story.append(NextPageTemplate('Normal'))
        Story.append(PageBreak())
        # empty paragraph just to assign the latitude = None to disable the header
        p = Paragraph('', self.style)
        Story.append(p)
        Story[-1].latitude = None
        p = Paragraph('<b>TABLE 4 - GHA Aries FOR THE YEARS '+str(self.epoch.tuple()[0]-4)+'-'+str(self.epoch.tuple()[0]+4)+'</b>', self.styleCentered)
        Story.append(p)
        Story.append(Spacer(0,inch/5))
        p = Paragraph('a. GHA Aries AT 00<super>h</super> ON THE FIRST DAY OF EACH MONTH', self.styleCentered)
        Story.append(p)
        Story.append(Spacer(0,inch/10))
        table, example_a = self.__table_gha_year()
        Story.append(table)
        Story.append(Spacer(0,inch/10))
        p = Paragraph('b. INCREMENT OF GHA Aries FOR DAYS AND HOURS', self.styleCentered)
        Story.append(p)
        Story.append(Spacer(0,inch/10))
        table, example_b = self.__table_gha_days()
        Story.append(table)
        Story.append(PageBreak())
        p = Paragraph('<b>TABLE 4 - GHA Aries FOR THE YEARS '+str(self.epoch-4)+'-'+str(self.epoch+4)+'</b>', self.styleCentered)
        Story.append(p)
        Story.append(Spacer(0,inch/5))
        p = Paragraph('c. INCREMENT OF GHA Aries FOR MINUTES AND SECONDS', self.styleCentered)
        Story.append(p)
        Story.append(Spacer(0,inch/10))
        table, example_c = self.__table_gha_minutes()
        Story.append(table)
        Story.append(Spacer(0,inch/10))
        text = '<i>Example.</i> The value of GHA Aries for '\
               +str(self.epoch.tuple()[0]+2)+\
               ' August 17 at 05<super>h</super> 11<super>m</super> 41<super>s</super> UT is (<b>a</b>) '\
               +self.__get_str_degrees(example_a)+\
               '+ (<b>b</b>) '\
               +self.__get_str_degrees(example_b)+\
               '+ (<b>c</b>) '\
               +self.__get_str_degrees(example_c)+\
               '= '+self.__get_str_degrees(ephem.degrees(example_a+example_b+example_c).norm)+'.'
        p = Paragraph(text, self.styleTable)
        Story.append(p)
        Story.append(PageBreak())
        return Story

    def __compute_polaris_data(self):
        polaris = all_stars['Polaris'].copy()
        polaris.compute(ephem.date(self.epoch), epoch = ephem.date(self.epoch))
        text = '<i>Polaris:</i> Mag.'\
           +str(round(polaris.mag,1))+\
           ', SHA '\
           +self.__get_str_degrees(2*pi-polaris.ra,1)+\
           ', DEC '\
           +self.__get_str_degrees(polaris.dec,1)
        return text

    def __compute_az_polaris(self):
        polaris = all_stars['Polaris'].copy()
        data = []
        # header of table
        header = []
        header2 = []
        header3 = []
        for block in range(2):
            header.append('LHA\nA')
            header.append('LATITUDE')
            for empty in range(6):
                header.append('')
            header2.append('')
            for lat in [0,30,50,55,60,65,70]:
                header2.append(str(lat)+'°')
        for columnn in range(16):
            header3.append('°')
        data.append(header)
        data.append(header2)
        data.append(header3)
        # body of table
        for lha in range(0,181,10):
            data_row = []
            for block in range(2):
                data_row.append(lha+(block*180))
                for lat in [0,30,50,55,60,65,70]:
                    actual_observer = self.__new_lha_zero_observer()
                    actual_observer.lon = actual_observer.lon+((lha+(block*180))*pi/180)
                    actual_observer.lat = lat*pi/180
                    polaris.compute(actual_observer)
                    az = round(polaris.az*180/pi,1)
                    if az == 360.0:
                        az = 0.0
                    data_row.append(az)
            data.append(data_row)
        style =[('FONT',(0,0),(-1,-1),'Helvetica', 6),
                ('FONT',(0,0),(-1,2),'Helvetica-Bold', 6),
                ('FONT',(0,3),(0,-1),'Helvetica-Bold', 6),
                ('FONT',(8,3),(8,-1),'Helvetica-Bold', 6),
                ('SPAN',(0,0),(0,1)),
                ('SPAN',(1,0),(7,0)),
                ('SPAN',(8,0),(8,1)),
                ('SPAN',(9,0),(15,0)),
                ('LINEAFTER',(0,0),(0,-1),0.2,colors.black),
                ('LINEAFTER',(2,1),(2,-1),0.2,colors.black),
                ('LINEAFTER',(5,1),(5,-1),0.2,colors.black),
                ('LINEAFTER',(7,0),(7,-1),2,colors.black),
                ('LINEAFTER',(8,0),(8,-1),0.2,colors.black),
                ('LINEAFTER',(10,1),(10,-1),0.2,colors.black),
                ('LINEAFTER',(13,1),(13,-1),0.2,colors.black),
                ('LINEBELOW',(0,1),(-1,1),0.2,colors.black),
                ('LEFTPADDING',(0,0),(-1,-1),6),
                ('RIGHTPADDING',(0,0),(-1,-1),6),
                ('BOTTOMPADDING',(0,0),(-1,-1),0),
                ('TOPPADDING',(0,0),(-1,-1),0),
                ('ALIGN', (0,0), (-1,1), 'CENTRE'),
                ('ALIGN', (0,2), (-1,-1), 'RIGHT'),
                ('VALIGN',(0,0), (-1,-1), 'MIDDLE'),
                ('BOX',(0,0),(-1,-1),2,colors.black)]
        t=Table(data, style = style,
                colWidths = self.PAGE_WIDTH/16*0.8,
                rowHeights = 8)
        return t

    def __compute_next_q_polaris(self, initial_lha, reverse = False):
        polaris = all_stars['Polaris'].copy()
        actual_observer = self.__new_lha_zero_observer()
        actual_observer.lon = actual_observer.lon+initial_lha
        polaris.compute(actual_observer)
        initial_q = round(-polaris.alt*10800/pi)
        while True:
            if not reverse:
                actual_observer.lon += (pi/10800)
            else:
                actual_observer.lon -= (pi/10800)
            polaris.compute(actual_observer)
            actual_q = round(-polaris.alt*10800/pi)
            lha_a = ephem.degrees(actual_observer.sidereal_time())
            if not actual_q == initial_q:
                return(lha_a,actual_q)

    def __compute_q_polaris(self):
        data = []
        temp_table = []
        # header of table
        header = []
        header2 = []
        header3 = []
        for block in range(8):
            header.append('LHA\nA')
            header.append('Q')
            header2.append('')
            header2.append('')
            header3.append('°    ´')
            header3.append('´')
        data.append(header)
        data.append(header2)
        data.append(header3)
        lha, q = self.__compute_next_q_polaris(ephem.degrees('0'),reverse = True)
        # for reverse search add 1 minute of arc to lha or the forward search will trip instantly
        lha += (pi/10800)
        for i in range(192):
            temp_table_row = []
            temp_table_row.append(self.__get_str_degrees(lha, signs = False))
            # get next lha/q
            lha, q = self.__compute_next_q_polaris(lha)
            if q > 0:
                q = '+'+str(q)
            temp_table_row.append(q)
            temp_table.append(temp_table_row)
        for row in range(24):
            data_row = []
            data_row2 = []
            for block in range(8):
                lha, q = temp_table[row+block*23]
                data_row.append(lha)
                data_row.append('')
                data_row2.append('')
                data_row2.append(q)
            data.append(data_row)
            if row < 23:
                data.append(data_row2)
            else:
                # insert empty row for padding
                data.append('')
        style =[('FONT',(0,0),(-1,-1),'Helvetica', 6),
                ('LINEBELOW',(0,1),(-1,1),0.2,colors.black),
                ('LEFTPADDING',(0,0),(-1,-1),6),
                ('RIGHTPADDING',(0,0),(-1,-1),6),
                ('BOTTOMPADDING',(0,0),(-1,-1),0),
                ('TOPPADDING',(0,0),(-1,-1),0),
                ('ALIGN', (0,0), (-1,1), 'CENTRE'),
                ('ALIGN', (0,2), (-1,-1), 'RIGHT'),
                ('VALIGN',(0,0), (-1,-1), 'MIDDLE'),
                ('BOX',(0,0),(-1,-1),2,colors.black)]
        for block in range(8):
            style.append(('SPAN',(block*2,0),(block*2,1)))
            style.append(('SPAN',(block*2+1,0),(block*2+1,1)))
            style.append(('LINEAFTER',(block*2+1,0),(block*2+1,-1),0.2,colors.black))
            style.append(('FONT',(block*2,3),(block*2,-1),'Helvetica-Bold', 6))

        t=Table(data, style = style,
                colWidths = self.PAGE_WIDTH/16*0.8,
                rowHeights = 8)
        return t

    def __compute_polaris_tables(self):
        Story = []
        Story.append(NextPageTemplate('Normal'))
        Story.append(PageBreak())
        # empty paragraph just to assign the latitude = None to disable the header
        p = Paragraph('', self.style)
        Story.append(p)
        Story[-1].latitude = None
        p = Paragraph('<b>TABLE 6 - CORRECTION (Q) FOR POLARIS</b>', self.styleCentered)
        Story.append(p)
        Story.append(Spacer(0,inch/5))
        table = self.__compute_q_polaris()
        Story.append(table)
        Story.append(Spacer(0,inch/10))
        text = 'The above table, which does not include refraction, gives the quantity Q to be applied to the corrected sextant altitude of\n<i>Polaris</i>'
        p = Paragraph(text, self.styleTable)
        Story.append(p)
        text = 'to give the latitude of the observer. In critical cases ascend.'
        p = Paragraph(text, self.styleTable)
        Story.append(p)
        text = self.__compute_polaris_data()
        p = Paragraph(text, self.styleTable)
        Story.append(p)        
        Story.append(Spacer(0,inch/5))
        p = Paragraph('<b>TABLE 7 - AZIMUTH OF POLARIS</b>', self.styleCentered)
        Story.append(p)
        Story.append(Spacer(0,inch/10))
        table = self.__compute_az_polaris()
        Story.append(table)
        Story.append(Spacer(0,inch/10))
        text = 'When Cassiopeia is left (right), <i>Polaris</i> is west (east).'
        p = Paragraph(text, self.styleTable)
        Story.append(p)
        return(Story)

        
    def go(self, fileName, progress_bar, lat_min, lat_max, epoch):
        self.epoch = epoch
        Story = []
        Story.extend(self.__compute_lat_table(progress_bar, lat_min, lat_max))
        Story.extend(self.__compute_gha_table())
        Story.extend(self.__compute_polaris_tables())

        if len(fileName ) > 0:
            try:
                with open(fileName, 'wb') as f:
                    doc = MyDocTemplate(f,
                                          pagesize = self.format,
                                          author= self.Author,
                                          subject = self.pageinfo,
                                          title = self.Title,
                                          #showBoundary=1,
                                          leftMargin=inch/3,
                                          rightMargin=inch/3,
                                          topMargin=inch/3,
                                          bottomMargin=inch/3)

                    #normal frame as for SimpleFlowDocument
                    frameT = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id='normal')

                    #Two Columns
                    frame1 = Frame(doc.leftMargin, doc.bottomMargin, doc.width/2-1, doc.height, id='col1')
                    frame2 = Frame(doc.leftMargin+doc.width/2+1, doc.bottomMargin, doc.width/2-1, doc.height, id='col2')

                    doc.addPageTemplates([PageTemplate(id='Title',frames=frameT,onPage=self.title),
                                          PageTemplate(id='Table',frames=[frame1,frame2],onPage=self.normalPagesFooter),
                                          PageTemplate(id='Normal',frames=frameT,onPage=self.normalPagesFooter)])
                    doc.build(Story)
                    if not progress_bar == None:
                        progress_bar['value'] = 100
            except Exception as e:
                messagebox.showinfo("Problem saving data", "Error: "+str(e))
                return False
        return True


class Stars(object):
    """
    Compute a stars Zn and Hc
    offtime is in minutes with early being negative
    """
    def __init__(self, precData, precomp, lower_limb = True):
        super().__init__()
        self.precomp = precomp
        self.__sd = 0.0
        basedate = self.precomp.observer.date.tuple()
        self.__date = ephem.Date((basedate[0], basedate[1], basedate[2], int(precData['time_hour']),int(precData['time_minute']), 0))
        offtime = int((self.__date - self.precomp.observer.date)/ephem.minute)
        if precData['name']=='':
            # if empy use Polaris (kind of messy to avoid extra work...)
            precData['name'] = 'Polaris'
        if precData['name']=='Moon':
            self.star = ephem.Moon()
        elif precData['name']=='Sun':
            self.star = ephem.Sun()
        elif precData['name']=='Venus':
            self.star = ephem.Venus()
        elif precData['name']=='Mars':
            self.star = ephem.Mars()        
        elif precData['name']=='Jupiter':
            self.star = ephem.Jupiter()        
        elif precData['name']=='Saturn':
            self.star = ephem.Saturn()        
        else:
            try:
                self.star = all_stars[precData['name']].copy()
            except KeyError:
                raise Exception("Unknown body or no name entered!")
        # precdata['sext'] is an integer (min of arc)
        self.sextant_error = ephem.degrees(precData['sext']/10800*pi)
        if precomp.acc_uso:
            temp_press = self.precomp.observer.pressure
            self.precomp.observer.pressure = 0
            self.star.compute(self.precomp.observer)
            self.precomp.observer.pressure = temp_press
            if lower_limb:
                self.__sd = self.star.radius
            else:
                self.__sd = -self.star.radius
            # calculate the LHA based on the LHA of Aries
            self.lha = ephem.degrees(2*pi-(self.precomp.lha_a-self.star.a_ra)).norm
            if self.star.name=='Polaris':
                self.__corr_hc = self.precomp.observer.lat
            else:
                self.__corr_hc = self.__get_hc()
            if self.star.name=='Moon':
                # or asin(sin(HP)*cos(Hc)), gives parallax error based on actual measured altitude
                self.__pa = ephem.degrees(asin(6371/(self.star.earth_distance * 149597870.7))*cos(self.__corr_hc))
            else:
                self.__pa = 0

            self.__zn = self.__get_zn()

            self.__moo = ephem.degrees(-offtime * self.__get_moo())
            self.__mob = ephem.degrees(-offtime * self.__get_mob())
            self.__refr = self.__get_refr()
            self.__total_adj = ephem.degrees(self.__moo + self.__mob + self.__sd + self.sextant_error + self.__refr + self.q + self.__pa)
            self.__adj_hc = ephem.degrees(self.__corr_hc - self.__total_adj)

        else:
            self.precomp.observer.pressure = 0
            if DEBUG.ON():
                self.precomp.observer.date = self.__date
            __observer_lat = self.precomp.observer.lat
            __observer_long = self.precomp.observer.long
            __observer_date = self.precomp.observer.date
            self.star.compute(self.precomp.observer)
            # calculate the LHA based on the LHA of Aries (used to compute Q for Polaris)
            self.lha = ephem.degrees(2*pi-(self.precomp.lha_a-self.star.a_ra)).norm
            self.__corr_hc = self.star.alt
            self.__zn = self.star.az

            # move observer position to account for motion of observer
            d = s2d(self.precomp.speed, offtime) # in radians
            (self.precomp.observer.lat, self.precomp.observer.lon) = bear_dist(self.precomp.observer.lat, self.precomp.observer.long, self.precomp.track, d)
            # now adjust time to account for motion of body
            
            self.precomp.observer.date = self.__date
            self.precomp.observer.compute_pressure()
            self.star.compute(self.precomp.observer)
            if lower_limb:
                self.__sd = self.star.radius
            else:
                self.__sd = -self.star.radius
            self.__adj_hc = ephem.degrees(self.star.alt-self.__sd-self.sextant_error)

            self.precomp.observer.lat = __observer_lat
            self.precomp.observer.long = __observer_long
            self.precomp.observer.date = __observer_date

            self.__moo = None
            self.__mob = None
            self.__refr = None
            self.__pa = None
            self.__total_adj = ephem.degrees(self.__corr_hc - self.__adj_hc)

        # compute postition of intercept along Zn
        if precData['ho_dd']!='' and precData['ho_mm']!='':
            self.intercept = 10800*(ephem.degrees(precData['ho_dd']+':'+precData['ho_mm'])-self.__adj_hc)/pi
        else:
            self.intercept = 0
        self.int_lat, self.int_lon = bear_dist(self.precomp.corrected_lat, self.precomp.corrected_lon, self.__zn, s2d(self.intercept))

    def __get_hc(self):
        return ephem.degrees(asin(sin(self.precomp.observer.lat)*sin(self.star.a_dec)+cos(self.lha)*cos(self.star.a_dec)*cos(self.precomp.observer.lat)))

    def __get_zn(self):
        z = acos((sin(self.star.a_dec)-sin(self.precomp.observer.lat)*sin(self.__corr_hc))/(cos(self.__corr_hc)*cos(self.precomp.observer.lat)))
        if self.precomp.observer.lat>0:
            if self.lha<pi:
                return ephem.degrees(z)
            else:
                return ephem.degrees(2*pi-z)
        else:
            if self.lha<pi:
                return ephem.degrees(pi-z)
            else:
                return ephem.degrees(pi+z)

    def __get_moo(self):
        """ Motion of observer for 1 minute """
        return ephem.degrees((self.precomp.speed/3600)*cos(self.precomp.track-self.__zn)*pi/180)

    def __get_mob(self):
        """ Motion of Body for 1 minute """
        return ephem.degrees(cos(self.precomp.observer.lat)*sin(self.__zn)*pi/720)

    def __get_refr(self):
        """ Compute refraction based on Hc """
        # good down to 15 deg Hc
        return ephem.degrees(-0.00452 * self.precomp.observer.pressure / ((273 + self.precomp.observer.temp) * tan(self.__corr_hc))/60)

    @property
    def q(self):
        if self.star.name=='Polaris':
            return ephem.degrees(self.__corr_hc - self.__get_hc())
        else:
            return 0

    @property
    def moo(self):
        return self.__moo

    @property
    def mob(self):
        return self.__mob

    @property
    def refr(self):
        return self.__refr

    @property
    def sd(self):
        return self.__sd

    @property
    def pa(self):
        return self.__pa

    @property
    def total_adj(self):
        return self.__total_adj

    @property
    def zn(self):
        return self.__zn
        
    @property
    def corr_hc(self):
        return self.__corr_hc
    
    @property
    def adj_hc(self):
        return self.__adj_hc

    @property
    def date(self):
        return self.__date

    @property
    def name(self):
        return self.star.name

    @property
    def gha(self):
        return ephem.degrees(self.precomp.gha-self.star.a_ra).norm

    @property
    def dec(self):
        return self.star.a_dec

class Precomputation(object):
    """
    Holds basic data to perform the precomputation
    """
    def __init__(self, precData):
        self.__dr_lat_r = ephem.degrees((precData['dr_lat_dd']+':'+precData['dr_lat_mm']))
        self.__dr_lon_r = ephem.degrees((precData['dr_lon_dd']+':'+precData['dr_lon_mm']))
        self.altitude = precData['alt']
        self.__track_r = ephem.degrees(precData['track']*pi/180)
        self.speed = precData['speed']
        self.epoch = precData['epoch']
        self.dtg = ephem.Date((precData['year'], precData['month'], precData['day'],
                               precData['hour'], precData['minute'], 0))
        self.free_running = precData['free_running']
        self.acc_uso = precData['acc_uso']

    def compute(self):
        # compute GHA
        greenwich = ephem.Observer()
        greenwich.epoch = self.epoch
        greenwich.date = self.dtg
        greenwich.long = '0:0:0'
        self.gha = ephem.degrees(greenwich.sidereal_time())
        # find nearest whole lat as assumed latitude
        self.assumed_lat = ephem.degrees(round(self.__dr_lat_r * 180 / pi) * pi / 180)
        if DEBUG.ON():
            self.assumed_lat = ephem.degrees(self.__dr_lat_r)
        # find nearest lon that will give a whole LHA
        self.assumed_lon = ephem.degrees(((round((self.__dr_lon_r + self.gha) * 180 / pi))* pi / 180) - self.gha)
        if DEBUG.ON():
            self.assumed_lon = ephem.degrees(self.__dr_lon_r)
        # calculate resulting LHA of Aries
        self.lha_a = ephem.degrees(self.assumed_lon + self.gha)
        # Calculate Coriolis and rhumb line effects and move the assumed to the corrected position (intercepts should be applied to this position)
        self.coriolis = self.get_coriolis()
        self.rhumbline = self.get_rhumbline()
        # correct assumed pos for coriolis and rhumb line
        self.corrected_lat, self.corrected_lon = bear_dist(self.assumed_lat, self.assumed_lon, ephem.degrees(self.__track_r+pi/2), s2d(self.coriolis+self.rhumbline))
        # set the observer to our assumed position and fix time
        self.observer = ephem.Observer()
        self.observer.date = self.dtg
        self.observer.epoch = self.epoch
        self.observer.lat = self.assumed_lat
        self.observer.lon = self.assumed_lon
        self.observer.elevation = self.altitude * 100 / 3.2808399
        self.observer.temp = 15 - (self.altitude / 5)
        self.observer.compute_pressure()
        
    def get_coriolis(self):
        """ Coriolis in NM """
        return 0.0265*self.speed*sin(self.assumed_lat)

    def get_rhumbline(self):
        """ Get rhumb line in NM """
        if not self.free_running:
            return 0.146*pow(self.speed/100,2)*sin(self.__track_r)*tan(self.assumed_lat)
        else:
            return 0.0
        
    @property
    def dr_lat(self):
        return self.__dr_lat_r

    @property
    def dr_lon(self):
        return self.__dr_lon_r
    
    @property
    def track(self):
        return self.__track_r

# a subclass of Canvas for dealing with resizing of windows
class ResizingCanvas(tk.Canvas):
    def __init__(self,parent,**kwargs):
        tk.Canvas.__init__(self,parent,bd=0,highlightthickness=0,**kwargs)
        tk.Canvas.configure(self,background="white")
        tk.Canvas.configure(self,cursor="cross_reverse")
        self.bind("<Configure>", self.configure)
        self.bind("<1>", self.OnMouseDown)
        self.cosphi = 1
        self.center_lon = 0
        self.center_lat = 100
        self.polygon_list = None
        self.fix_coords = None
        self.precomp = None
        self.intercepts = {}

    def OnMouseDown(self, event):
        lat, lon = self.__p2w((event.x, event.y))
        self.delete('mouse')
        # do not draw text unless there is a grid
        if not self.center_lat==100:
            self.create_oval(event.x, event.y, event.x, event.y, width=2, tags= 'mouse')
            text_lat = self.create_text(event.x+10, event.y+10, anchor = 'nw' , tags = 'mouse', text = str(ephem.degrees(lat)))
            box_lat = self.create_rectangle(self.bbox(text_lat), fill = 'white', outline = '', tags = 'mouse')
            text_lon = self.create_text(event.x+10, event.y+25, anchor = 'nw' , tags = 'mouse', text = str(ephem.degrees(lon)))
            box_lon = self.create_rectangle(self.bbox(text_lon), fill = 'white', outline = '', tags = 'mouse')
            self.tag_raise(text_lat, box_lat)
            self.tag_raise(text_lon, box_lon)

    def configure(self,event):
        if self.center_lat==100:
            self.delete("all")
            w, h = event.width, event.height
            xy = 0, 0, w-1, h-1
            self.create_rectangle(xy)
            self.create_line(xy, fill = 'darkgrey')
            xy = w-1, 0, 0, h-1
            self.create_line(xy, fill = 'darkgrey')
            self.create_text(w/2, h/2, text = 'Calculate to display map')
        else:
            self.draw(self.precomp)
        if self.polygon_list:
            self.polygon(self.polygon_list)
        if self.fix_coords:
            self.fix(self.fix_coords)
        try:
            for number, (lat, lon, track, intercept) in self.intercepts.items():
                self.intercept(number, lat, lon, track, intercept)
        except:
            pass

    def draw(self, precomp):
        self.delete("all")
        if self.precomp!=precomp:
            self.precomp = precomp
            self.polygon_list = None
            self.fix_coords = None
            self.intercepts = {}
        self.precomp = precomp
        self.__draw_grid()
        self.circle(self.precomp.dr_lat, self.precomp.dr_lon, outline='black')
        self.circle(self.precomp.assumed_lat, self.precomp.assumed_lon, outline='orange')
        self.circle(self.precomp.corrected_lat, self.precomp.corrected_lon, outline='red')

    def intercept(self, number, lat, lon, track, intercept = None):
        self.intercepts[number]=((lat, lon, track, intercept))
        dash = ''
        for x in range(number):
            dash += '5 5 '
        dash += '15 5'
        """ draw a line along track for dist NM - WARNING: dist in NM!!! """
        if intercept:
            self.create_line(self.__w2p((lat, lon)), self.__w2p(bear_dist(lat, lon, track, s2d(intercept))), dash=dash)
            int_lat, int_lon = bear_dist(lat,lon, track, s2d(intercept))
            self.create_line(self.__w2p((int_lat, int_lon)), self.__w2p(bear_dist(int_lat, int_lon, track+pi/2, s2d(60))))
            self.create_line(self.__w2p((int_lat, int_lon)), self.__w2p(bear_dist(int_lat, int_lon, track-pi/2, s2d(60))))
        else:
            # default to 60 NM
            dist_rad = pi/180
            self.create_line(self.__w2p((lat, lon)), self.__w2p(bear_dist(lat, lon, track, s2d(-60))), dash=dash)
            return self.create_line(self.__w2p((lat, lon)), self.__w2p(bear_dist(lat, lon, track, s2d(60))), dash=dash)

    def circle(self, lat, lon, outline = 'black'):
        x, y = self.__w2p((lat,lon))
        return self.create_oval(x, y, x, y, width=3, outline=outline)

    def polygon(self, list):
        self.polygon_list = list
        points = []
        for x in list:
            points.append(self.__w2p(x))
        return self.create_polygon(points, fill = '', width=2, outline='red')

        x, y = self.__w2p((lat,lon))
        return self.create_oval(x, y, x, y, width=3, outline='black')

    def fix(self, latlon):
        self.fix_coords = latlon
        lat, lon = latlon
        # Dreieck mit Punkt fuer Fix zeichnen
        posx, posy = self.__w2p((lat, lon))
        size = 5
        points = [posx, posy-size, posx+size, posy+size, posx-size, posy+size, posx, posy-size]
        self.create_polygon(points, fill = '', width=2, outline='red')
        self.create_oval(posx-1, posy-1, posx+1, posy+1)

    def __w2p(self, coord):
        delta_lat = (coord[0] - self.center_lat)*180/pi
        delta_lon = (coord[1] - self.center_lon)*180/pi
        x = self.winfo_width()/2 + delta_lon*2/3*self.cosphi*self.winfo_height()
        y = self.winfo_height()/2 - delta_lat*2/3*self.winfo_height()
        return x,y      

    def __p2w(self, coord):
        delta_lat = (self.winfo_height()/2 - coord[1])*3/(2*self.winfo_height())
        delta_lon = (coord[0] - self.winfo_width()/2)*3/(2*self.cosphi*self.winfo_height())
        lat = ephem.degrees((delta_lat*pi/180) + self.center_lat)
        lon = ephem.degrees((delta_lon*pi/180) + self.center_lon)
        return lat, lon  

    def __draw_grid(self):
        plotheight = self.winfo_height()
        plotwidth = self.winfo_width()
        # departure
        self.cosphi = cos(self.precomp.assumed_lat)
        # define center of plot
        self.center_lon = ephem.degrees(round(self.precomp.assumed_lon/0.00872664626)*0.00872664626)
        self.center_lat = ephem.degrees(round(self.precomp.assumed_lat/0.00872664626)*0.00872664626)
        # LAT
        delta = 1/3*plotheight
        x = 0
        number = 0
        while x < plotheight/2:
            self.create_text(40, plotheight/2 - x - 10, text = str(ephem.degrees(self.center_lat+number*30*ephem.arcminute)))
            self.create_line(0, plotheight/2 - x, plotwidth, plotheight/2 - x, fill = 'darkgrey')
            self.create_text(40, plotheight/2 + x - 10, text = str(ephem.degrees(self.center_lat-number*30*ephem.arcminute)))
            self.create_line(0, plotheight/2 + x, plotwidth, plotheight/2 + x, fill = 'darkgrey')
            x += delta
            number += 1
        # LON, account for departure (Abweitung)
        delta = 1/3*self.cosphi*plotheight
        x = 0
        number = 0
        while x < plotwidth/2:
            self.create_text(plotwidth/2 - x, 5, text = str(ephem.degrees(self.center_lon-number*30*ephem.arcminute)))
            self.create_line(plotwidth/2 - x, 0, plotwidth/2 - x, plotheight, fill = 'darkgrey')
            self.create_text(plotwidth/2 + x, 5, text = str(ephem.degrees(self.center_lon+number*30*ephem.arcminute)))
            self.create_line(plotwidth/2 + x, 0, plotwidth/2 + x, plotheight, fill = 'darkgrey')
            x += delta
            number += 1


class Gui(tk.Frame):
    """ Displays GUI and starts computation """
    def __init__(self, master = None):
        tk.Frame.__init__(self, master)
        self.master = master
        self.master.wm_title("Celestial Precomputation © Lars Neumann")
        #start with the window maximised
        #master.state('zoomed')
        # intercept system call to close the window
        self.master.protocol('WM_DELETE_WINDOW', self.quit_gui)

        self.grid(row=0, sticky = tk.W+tk.E+tk.N+tk.S)

        self.precData = {}
        self.time = datetime.utcnow()
        self.invaliddata = {}

        self.__defaults()
        
        # create a pulldown menu, and add it to the menu bar
        menubar = tk.Menu(master)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Load...", command=self.load)
        filemenu.add_command(label="Save...", command=self.save)
        filemenu.add_separator()
        filemenu.add_command(label="Exit...", command=self.quit_gui)
        menubar.add_cascade(label="File", menu=filemenu)
        menubar.add_command(label="Clear", command=self.clear)
        menubar.add_command(label="Generate PDF...", command=self.generate_pdf_window)
        debugmenu = tk.Menu(menubar, tearoff=0)
        self.debug = tk.BooleanVar()
        debugmenu.add_checkbutton(label="On", variable=self.debug)
        menubar.add_cascade(label="Debug", menu=debugmenu)
        menubar.add_command(label="About", command=self.aboutwindow)

        # display the menu
        master.config(menu=menubar)
        
        frame_date = tk.LabelFrame(self, text="Fix Time", padx=5, pady=5)
        frame_date.grid(sticky= tk.W+tk.E+tk.N+tk.S)
        
        # valid percent substitutions (from the Tk entry man page)
        # %d = Type of action (1=insert, 0=delete, -1 for others)
        # %i = index of char string to be inserted/deleted, or -1
        # %P = value of the entry if the edit is allowed
        # %s = value of entry prior to editing
        # %S = the text string being inserted or deleted, if any
        # %v = the type of validation that is currently set
        # %V = the type of validation that triggered the callback
        #      (key, focusin, focusout, forced)
        # %W = the tk name of the widget
        vcmdyear = (master.register(self.validate),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W', 'year')
        vcmdyear_epoch = (master.register(self.validate),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W', 'year_epoch')
        vcmdmonth = (master.register(self.validate),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W', 'month')
        vcmdday = (master.register(self.validate),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W', 'day')
        vcmdhour = (master.register(self.validate),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W', 'hour')
        vcmdlat = (master.register(self.validate),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W', 'lat')
        vcmdlon = (master.register(self.validate),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W', 'lon')
        vcmdho_dd = (master.register(self.validate),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W', 'ho_dd')
        vcmdho_mm = (master.register(self.validate),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W', 'ho_mm')
        vcmdbase60 = (master.register(self.validate),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W', 'base60')
        vcmdflgs = (master.register(self.validate),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W', 'flgs')
        vcmdbase360 = (master.register(self.validate),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W', 'base360')
        vcmd_sext = (master.register(self.validate),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W', 'sext')

        self.day = self.makeentry(frame_date, "Date", 2, validate='all', vcmd=vcmdday)
        self.month = self.makeentry(frame_date, "", 2, rightof=self.day, validate='all', vcmd=vcmdmonth)
        self.year = self.makeentry(frame_date, "", 4, rightof=self.month, validate='all', vcmd=vcmdyear)
        self.hour = self.makeentry(frame_date, "Fix Time", 2, validate='all', vcmd=vcmdhour)
        self.minute = self.makeentry(frame_date, "", 2, rightof=self.hour, validate='all', vcmd=vcmdbase60)

        # button to automatically calculate planned mid times from fix time
        self.std_plan_btn = tk.Button(frame_date, text="Std plan", command=self.stdPlan)
        self.std_plan_btn.grid(sticky=tk.W+tk.E+tk.N+tk.S, columnspan = 4, padx = 5, pady = 5)

        frame_owndata = tk.LabelFrame(self, text="Own Data", padx=5, pady=5)
        frame_owndata.grid(sticky=tk.W+tk.E+tk.N+tk.S)

        self.alt = self.makeentry(frame_owndata, "FL", 3, validate='all', vcmd=vcmdflgs)
        self.track = self.makeentry(frame_owndata, "TC", 3, validate='all', vcmd=vcmdbase360)
        self.speed = self.makeentry(frame_owndata, "GS", 3, validate='all', vcmd=vcmdflgs)
        self.dr_lat_dd = self.makeentry(frame_owndata, "DR Lat", 3, validate='all', vcmd=vcmdlat)
        self.dr_lat_mm = self.makeentry(frame_owndata, "", 2, rightof=self.dr_lat_dd, validate='all', vcmd=vcmdbase60)
        self.dr_lon_dd = self.makeentry(frame_owndata, "DR Lon", 4, validate='all', vcmd=vcmdlon)
        self.dr_lon_mm = self.makeentry(frame_owndata, "", 2, rightof=self.dr_lon_dd, validate='all', vcmd=vcmdbase60)
        self.epoch = self.makeentry(frame_owndata, "Epoch", 4, validate='all', vcmd=vcmdyear_epoch)

        frame_bodies = tk.LabelFrame(self, text="Bodies", padx=5, pady=5)
        frame_bodies.grid(sticky=tk.W+tk.E+tk.N+tk.S)

        self.body1_name = self.makeentry(frame_bodies, "Body 1", 15, columnspan=3)
        self.body1_pl_time_hour = self.makeentry(frame_bodies, "Pl. mid-time", 2, column=1, validate='all', vcmd=vcmdhour)
        self.body1_pl_time_minute = self.makeentry(frame_bodies, "", 2, rightof=self.body1_pl_time_hour, validate='all', vcmd=vcmdbase60)
        self.body1_ho_dd = self.makeentry(frame_bodies, "Ho", 2, column=1, validate='all', vcmd=vcmdho_dd)
        self.body1_ho_mm = self.makeentry(frame_bodies, "", 2, rightof=self.body1_ho_dd, validate='all', vcmd=vcmdho_mm)
        self.body1_sext = self.makeentry(frame_bodies, "Pers/Sext", 3, column=1, columnspan=2, validate='all', vcmd=vcmd_sext)
        self.body2_name = self.makeentry(frame_bodies, "Body 2", 15, columnspan=3)
        self.body2_pl_time_hour = self.makeentry(frame_bodies, "Pl. mid-time", 2, column=1, validate='all', vcmd=vcmdhour)
        self.body2_pl_time_minute = self.makeentry(frame_bodies, "", 2, rightof=self.body2_pl_time_hour, validate='all', vcmd=vcmdbase60)
        self.body2_ho_dd = self.makeentry(frame_bodies, "Ho", 2, column=1, validate='all', vcmd=vcmdho_dd)
        self.body2_ho_mm = self.makeentry(frame_bodies, "", 2, rightof=self.body2_ho_dd, validate='all', vcmd=vcmdho_mm)
        self.body2_sext = self.makeentry(frame_bodies, "Pers/Sext", 3, column=1, columnspan=2, validate='all', vcmd=vcmd_sext)
        self.body3_name = self.makeentry(frame_bodies, "Body 3", 15, columnspan=3)
        self.body3_pl_time_hour = self.makeentry(frame_bodies, "Pl. mid-time", 2, column=1, validate='all', vcmd=vcmdhour)
        self.body3_pl_time_minute = self.makeentry(frame_bodies, "", 2, rightof=self.body3_pl_time_hour, validate='all', vcmd=vcmdbase60)
        self.body3_ho_dd = self.makeentry(frame_bodies, "Ho", 2, column=1, validate='all', vcmd=vcmdho_dd)
        self.body3_ho_mm = self.makeentry(frame_bodies, "", 2, rightof=self.body3_ho_dd, validate='all', vcmd=vcmdho_mm)
        self.body3_sext = self.makeentry(frame_bodies, "Pers/Sext", 3, column=1, columnspan=2, validate='all', vcmd=vcmd_sext)
        
        self.acc_uso_var = tk.BooleanVar()
        self.acc_uso = tk.Checkbutton(self, text="Calculation acc. USO", var = self.acc_uso_var, offvalue=False, onvalue=True)
        self.acc_uso.grid(sticky=tk.W)
        self.free_running_var = tk.BooleanVar()
        self.free_running = tk.Checkbutton(self, text="Free running gyro", var = self.free_running_var, offvalue=False, onvalue=True)
        self.free_running.grid(sticky=tk.W)
        self.lower_limb_var = tk.BooleanVar()
        self.lower_limb = tk.Checkbutton(self, text="Lower limb", var = self.lower_limb_var, offvalue=False, onvalue=True)
        self.lower_limb.grid(sticky=tk.W)

        frame_canvas = tk.LabelFrame(self, text="Fix", padx=5, pady=5)
        frame_canvas.grid(sticky=tk.W+tk.E+tk.N+tk.S, row=0, column=1, rowspan=5)
        # make the plot resizeable
        master.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        frame_canvas.columnconfigure(0, weight=1, minsize=300)
        master.rowconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        frame_canvas.rowconfigure(0, weight=1, minsize=300)

        self.plot = ResizingCanvas(frame_canvas, width=600, height=600)
        self.plot.grid(sticky=tk.W+tk.E+tk.N+tk.S)
        
        frame_output = tk.LabelFrame(self, text="Output", padx=5, pady=5)
        frame_output.grid(row=0, column=2, rowspan=5, sticky=tk.N+tk.W+tk.S)

        self.gha_out = self.makeoutput(frame_output, 'GHA Aries:', width=11)
        self.assumed_lon_out = self.makeoutput(frame_output, 'Assumed LONG:', width=11)
        self.lha_out = self.makeoutput(frame_output, 'LHA Aries:', width=11)
        self.assumed_lat_out = self.makeoutput(frame_output, 'Assumed LAT:', width=11)
        self.coriolis = self.makeoutput(frame_output, 'Coriolis:', width=11)
        self.rhumbline = self.makeoutput(frame_output, 'Rhumb Line:', width=11)
        self.corrected_lat_out, lat_out = self.makeoutput(frame_output, 'Resulting Pos.:', width=11, returnlabel=True)
        self.corrected_lon_out = self.makeoutput(frame_output, '', width=11, rightof=lat_out)
        self.body = ({},{},{})
        fields = OrderedDict([('name','Body:'), ('gha','GHA:'), ('dec','DEC:'), ('date','Planned mid-time:'), ('chc','Corr. Hc:'),
                              ('moo','Motion Obs.:'), ('mob','Motion Body:'), ('refr','Refraction:'), ('sext','Pers./Sext:'),
                              ('pa','Moon PA:'), ('sd','Sun/Moon SD:'), ('p_q','Polaris Q:'),
                              ('tadj','Total Adj.:'), ('ahc','Adjusted Hc:'), ('int','Intercept:'), ('zn','Zn:')])
        for field,text in fields.items():
            for number in range(3):
                if number==0:
                    self.body[number][field], lastfield = self.makeoutput(frame_output, text, width=11, returnlabel=True)
                elif number<2:
                    self.body[number][field], lastfield = self.makeoutput(frame_output, '', rightof=lastfield, width=11, returnlabel=True)
                else:
                    self.body[number][field] = self.makeoutput(frame_output, '', rightof=lastfield, width=11)

        self.mpp_lat_out, mpp_lat = self.makeoutput(frame_output, 'MPP.:', width=11, returnlabel=True)
        self.mpp_lon_out = self.makeoutput(frame_output, '', width=11, rightof=mpp_lat)

        self.delta_out = self.makeoutput(frame_output, 'Delta (NM):', width=11)

        self.output = tk.Message(frame_output, width=300)
        self.output.grid(columnspan=4, sticky=tk.W+tk.E+tk.N+tk.S)

        frame_buttons = tk.Frame(master)
        frame_buttons.grid()

        self.quit_btn = tk.Button(frame_buttons, text="QUIT", fg="red", command=self.quit_gui)
        self.quit_btn.grid(row=0, padx = 5, pady = 5)

        self.clear_btn = tk.Button(frame_buttons, text="Clear", command=self.clear)
        self.clear_btn.grid(row=0, column=1, padx = 5, pady = 5)

        self.save_btn = tk.Button(frame_buttons, text="Save", command=self.save)
        self.save_btn.grid(row=0, column=2, padx = 5, pady = 5)
        self.load_btn = tk.Button(frame_buttons, text="Load", command=self.load)
        self.load_btn.grid(row=0, column=3, padx = 5, pady = 5)
        
        self.calculate_btn = tk.Button(frame_buttons, text="Calculate", command=self.calculate)
        self.calculate_btn.grid(row=0, column=4, padx = 5, pady = 5)
        master.bind("<Key>", self.valid)
        master.update()
        # do not allow the window to be smaller than after start
        master.minsize(master.winfo_width()-300, master.winfo_height()-100)

        # check if default.cel exists and load file
        if os.path.isfile('default.cel'):
            self.load(default = True)
        else:
            self.initialize()

    def valid(self, event = None):
        try:
            self.__gettime()
            if self.invaliddata['year']:
                raise ValueError
        except:
            self.std_plan_btn.config(state=tk.DISABLED)
        else:
            self.std_plan_btn.config(state=tk.NORMAL)
        if sum(self.invaliddata.values()):
            self.calculate_btn.config(state=tk.DISABLED)
            self.save_btn.config(state=tk.DISABLED)
        else:
            self.calculate_btn.config(state=tk.NORMAL)
            self.save_btn.config(state=tk.NORMAL)

    def clear(self):
        self.__defaults()
        fixtime = self.__default_fixtime()
        self.__plan_time(fixtime)
        self.initialize()

    def stdPlan(self):
        fixtime = datetime(self.precData['year'], self.precData['month'], self.precData['day'],
                       self.precData['hour'], self.precData['minute'], 0)
        self.__plan_time(fixtime)
        self.initialize()

    def __default_fixtime(self):
        # get current time in UTC, then find the next minute dividable by 4
        # that is at least 16 minutes in the future
        self.time = datetime.utcnow()
        fixtime = self.time
        next4minute = roundto(fixtime.time().minute)
        fixtime = fixtime.replace(minute = next4minute, second = 0)
        while fixtime-self.time < timedelta(minutes=16):
            fixtime+=timedelta(minutes=4)
        self.precData['year'] = fixtime.strftime('%Y')
        self.precData['month'] = fixtime.strftime('%m').lstrip('0')
        self.precData['day'] = fixtime.strftime('%d').lstrip('0')
        if fixtime.strftime('%H')!='00':
            self.precData['hour'] = fixtime.strftime('%H').lstrip('0')
        else:
            self.precData['hour'] = '0'
        if fixtime.strftime('%M')!='00':
            self.precData['minute'] = fixtime.strftime('%M').lstrip('0')
        else:
            self.precData['minute'] = '0'
        return fixtime

    def __plan_time(self, fixtime):
        # insert planned mid times that are 4 minutes apart starting 4 minutes prior fix time
        timebody3 = fixtime-timedelta(minutes=4)
        timebody2 = timebody3-timedelta(minutes=4)
        timebody1 = timebody2-timedelta(minutes=4)
        self.precData['body1']['time_hour'] = timebody1.strftime('%H')
        self.precData['body1']['time_minute'] = timebody1.strftime('%M')
        self.precData['body2']['time_hour'] = timebody2.strftime('%H')
        self.precData['body2']['time_minute'] = timebody2.strftime('%M')
        self.precData['body3']['time_hour'] = timebody3.strftime('%H')
        self.precData['body3']['time_minute'] = timebody3.strftime('%M')        

    def __defaults(self):
        self.precData['dr_lat_dd'] = '53'
        self.precData['dr_lat_mm'] = '00'
        self.precData['dr_lon_dd'] = '008'
        self.precData['dr_lon_mm'] = '00'
        self.precData['alt'] = '0'
        self.precData['track'] = '0'
        self.precData['speed'] = '0'
        self.precData['acc_uso'] = True
        self.precData['epoch'] = '2010'
        self.precData['free_running'] = False
        self.precData['body1'] = {}
        self.precData['body1']['name'] = ''
        self.precData['body1']['ho_dd'] = ''
        self.precData['body1']['ho_mm'] = ''
        self.precData['body1']['sext'] = '-12'
        self.precData['body2'] = {}
        self.precData['body2']['name'] = ''
        self.precData['body2']['ho_dd'] = ''
        self.precData['body2']['ho_mm'] = ''
        self.precData['body2']['sext'] = '-12'
        self.precData['body3'] = {}
        self.precData['body3']['name'] = ''
        self.precData['body3']['ho_dd'] = ''
        self.precData['body3']['ho_mm'] = ''
        self.precData['body3']['sext'] = '-12'

    def initialize(self):
        self.year.delete(0, tk.END)
        self.year.insert(0, self.precData['year'])
        self.month.delete(0, tk.END)
        self.month.insert(0, self.precData['month'])
        self.day.delete(0, tk.END)
        self.day.insert(0, self.precData['day'])
        self.hour.delete(0, tk.END)
        self.hour.insert(0, self.precData['hour'])
        self.minute.delete(0, tk.END)
        self.minute.insert(0, self.precData['minute'])
        self.dr_lat_dd.delete(0, tk.END)
        self.dr_lat_dd.insert(0, self.precData['dr_lat_dd'])
        self.dr_lat_mm.delete(0, tk.END)
        self.dr_lat_mm.insert(0, self.precData['dr_lat_mm'])
        self.dr_lon_dd.delete(0, tk.END)
        self.dr_lon_dd.insert(0, self.precData['dr_lon_dd'])
        self.dr_lon_mm.delete(0, tk.END)
        self.dr_lon_mm.insert(0, self.precData['dr_lon_mm'])
        self.alt.delete(0, tk.END)
        self.alt.insert(0, self.precData['alt'])
        self.track.delete(0, tk.END)
        self.track.insert(0, self.precData['track'])
        self.speed.delete(0, tk.END)
        self.speed.insert(0, self.precData['speed'])
        self.epoch.delete(0, tk.END)
        self.epoch.insert(0, self.precData['epoch'])
        if self.precData['acc_uso']:
            self.acc_uso.select()
        else:
            self.acc_uso.deselect()
        if self.precData['free_running']:
            self.free_running.select()
        else:
            self.free_running.deselect()
        if self.precData['lower_limb']:
            self.lower_limb.select()
        else:
            self.lower_limb.deselect()

        self.body1_name.delete(0, tk.END)
        self.body1_name.insert(0, self.precData['body1']['name'])
        self.body1_pl_time_hour.delete(0, tk.END)
        self.body1_pl_time_hour.insert(0, self.precData['body1']['time_hour'])
        self.body1_pl_time_minute.delete(0, tk.END)
        self.body1_pl_time_minute.insert(0, self.precData['body1']['time_minute'])
        self.body1_ho_dd.delete(0, tk.END)
        self.body1_ho_dd.insert(0, self.precData['body1']['ho_dd'])
        self.body1_ho_mm.delete(0, tk.END)
        self.body1_ho_mm.insert(0, self.precData['body1']['ho_mm'])
        self.body1_sext.delete(0, tk.END)
        self.body1_sext.insert(0, self.precData['body1']['sext'])
        self.body2_name.delete(0, tk.END)
        self.body2_name.insert(0, self.precData['body2']['name'])
        self.body2_pl_time_hour.delete(0, tk.END)
        self.body2_pl_time_hour.insert(0, self.precData['body2']['time_hour'])
        self.body2_pl_time_minute.delete(0, tk.END)
        self.body2_pl_time_minute.insert(0, self.precData['body2']['time_minute'])
        self.body2_ho_dd.delete(0, tk.END)
        self.body2_ho_dd.insert(0, self.precData['body2']['ho_dd'])
        self.body2_ho_mm.delete(0, tk.END)
        self.body2_ho_mm.insert(0, self.precData['body2']['ho_mm'])
        self.body2_sext.delete(0, tk.END)
        self.body2_sext.insert(0, self.precData['body2']['sext'])
        self.body3_name.delete(0, tk.END)
        self.body3_name.insert(0, self.precData['body3']['name'])
        self.body3_pl_time_hour.delete(0, tk.END)
        self.body3_pl_time_hour.insert(0, self.precData['body3']['time_hour'])
        self.body3_pl_time_minute.delete(0, tk.END)
        self.body3_pl_time_minute.insert(0, self.precData['body3']['time_minute'])
        self.body3_ho_dd.delete(0, tk.END)
        self.body3_ho_dd.insert(0, self.precData['body3']['ho_dd'])
        self.body3_ho_mm.delete(0, tk.END)
        self.body3_ho_mm.insert(0, self.precData['body3']['ho_mm'])
        self.body3_sext.delete(0, tk.END)
        self.body3_sext.insert(0, self.precData['body3']['sext'])
        # manually check if valid (i.e. on load, clear, plan) as no key was pressed
        self.valid()

    def __gettime(self):
        self.precData['year'] = int(self.year.get())
        self.precData['month'] = int(self.month.get())
        self.precData['day'] = int(self.day.get())
        self.precData['hour'] = int(self.hour.get())
        self.precData['minute'] = int(self.minute.get())

    def get(self):
        self.__gettime()
        self.precData['year'] = int(self.year.get())
        self.precData['month'] = int(self.month.get())
        self.precData['day'] = int(self.day.get())
        self.precData['hour'] = int(self.hour.get())
        self.precData['minute'] = int(self.minute.get())
        self.precData['dr_lat_dd'] = self.dr_lat_dd.get()
        self.precData['dr_lat_mm'] = self.dr_lat_mm.get()
        self.precData['dr_lon_dd'] = self.dr_lon_dd.get()
        self.precData['dr_lon_mm'] = self.dr_lon_mm.get()
        self.precData['alt'] = int(self.alt.get())
        self.precData['track'] = int(self.track.get())
        self.precData['speed'] = int(self.speed.get())
        self.precData['acc_uso'] = self.acc_uso_var.get()
        self.precData['lower_limb'] = self.lower_limb_var.get()
        self.precData['epoch'] = self.epoch.get()
        self.precData['free_running'] = self.free_running_var.get()
        self.precData['body1']['name'] = self.body1_name.get()
        self.precData['body1']['time_hour'] = self.body1_pl_time_hour.get()
        self.precData['body1']['time_minute'] = self.body1_pl_time_minute.get()
        self.precData['body1']['ho_dd'] = self.body1_ho_dd.get()
        self.precData['body1']['ho_mm'] = self.body1_ho_mm.get()
        self.precData['body1']['sext'] = int(self.body1_sext.get())
        self.precData['body2']['name'] = self.body2_name.get()
        self.precData['body2']['time_hour'] = self.body2_pl_time_hour.get()
        self.precData['body2']['time_minute'] = self.body2_pl_time_minute.get()
        self.precData['body2']['ho_dd'] = self.body2_ho_dd.get()
        self.precData['body2']['ho_mm'] = self.body2_ho_mm.get()
        self.precData['body2']['sext'] = int(self.body2_sext.get())
        self.precData['body3']['name'] = self.body3_name.get()
        self.precData['body3']['time_hour'] = self.body3_pl_time_hour.get()
        self.precData['body3']['time_minute'] = self.body3_pl_time_minute.get()
        self.precData['body3']['ho_dd'] = self.body3_ho_dd.get()
        self.precData['body3']['ho_mm'] = self.body3_ho_mm.get()
        self.precData['body3']['sext'] = int(self.body3_sext.get())
        # set debugging mode
        DEBUG.set(self.debug.get())


    def quit_gui(self):
        if messagebox.askokcancel('Close', 'Are you sure you want to close?'):
            self.master.destroy()

    def save(self):
        fileName = filedialog.asksaveasfilename(defaultextension = '.cel', filetypes = [("Celestial navigation dataset",".cel")])
        if len(fileName ) > 0:
            try:
                with open(fileName, 'wb') as f:
                    self.get()
                    # Pickle the 'data' dictionary using the highest protocol available.
                    pickle.dump(self.precData, f, pickle.HIGHEST_PROTOCOL)
            except Exception as e:
                messagebox.showinfo("Problem saving data", "Error: "+str(e))

    def load(self, default = False):
        if default:
            fileName = 'default.cel'
        else:
            fileName = filedialog.askopenfilename(defaultextension = '.cel', filetypes = [("Celestial navigation dataset",".cel")])
        if len(fileName ) > 0:
            try:
                with open(fileName, 'rb') as f:
                    # The protocol version used is detected automatically, so we do not
                    # have to specify it.
                    self.precData = pickle.load(f)
                    self.initialize()
            except:
                pass

    def generate_pdf_window(self):
        self.top = tk.Toplevel()
        self.top.title("Generate PDF file")
        self.top.resizable(width=False, height=False)
        self.top.focus_set()

        vcmdyear = (self.top.register(self.validate),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W', 'year_epoch')
        vcmdlat = (self.top.register(self.validate),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W', 'lat')

        textframe = tk.Frame(self.top)
        textframe.grid(row=0, sticky = tk.W+tk.E, padx = 5, pady = 5)
        dataframe = tk.Frame(self.top)
        dataframe.grid(row=1, padx = 5, pady = 5)
        buttonframe = tk.Frame(self.top)
        buttonframe.grid(row=2, padx = 5, pady = 5)

        msg = tk.Message(textframe, width = 320, text='You can generate a PDF file that is similar to the tables in Pub 249 Vol 1. Seven stars are selected for each block - the selection algorithm is different to the one the NGA uses so expect a somewhat different set of stars.')
        msg.grid(row = 0, sticky = tk.E+tk.W)
        msg2 = tk.Message(textframe, width = 320, text='This operation might take a while.')
        msg2.grid(row = 1, sticky = tk.E+tk.W)
        msg3 = tk.Message(textframe, width = 320, text='This window will close after the file has been created.')
        msg3.grid(row = 2, sticky = tk.E+tk.W)

        self.filename = self.makeoutput(dataframe, 'File name:', width = 40, row = 0)
        lat_min = self.makeentry(dataframe, 'Southern limit:', sticky = tk.W, width = 3, row = 1, validate = 'all', vcmd=vcmdlat)
        lat_max = self.makeentry(dataframe, 'Northern limit:', sticky = tk.W, width = 3, row = 2)
        lat_min.insert(0, '-89')
        lat_max.insert(0, '89')
        self.epoch = self.makeentry(dataframe, 'Epoch:', sticky = tk.W, width = 4, row = 3, validate = 'all', vcmd=vcmdyear)
        self.epoch.insert(0, datetime.now().year)
        self.filename.set('PUB249_Vol1_'+self.epoch.get()+'.pdf')

        progress = tk.IntVar()
        progress_bar = Progressbar(dataframe, length = 300)
        progress_bar.grid(row = 4, columnspan = 2, pady = 5)

        dismissbutton = tk.Button(buttonframe, text="Dismiss", command=self.top.destroy)
        filebutton = tk.Button(buttonframe, text="Change File Name", command=self.askpdffilename)
        startbutton = tk.Button(buttonframe, text="Start",
                                command=lambda:self.make_pdf(progress_bar, int(lat_min.get()), int(lat_max.get())))
        dismissbutton.grid(row = 0, padx = 5)
        filebutton.grid(row = 0, column = 2, padx = 5)
        startbutton.grid(row = 0, column = 3, padx = 5)

    def make_pdf(self, progress_bar, lat_min, lat_max):
        # check for valid input, basic validation has been performed with vcmd
        if lat_min > lat_max:
            messagebox.showerror(
            "Check input", "Southern limit must be to the south of the northern limit...")
            self.top.lift()
        if not 1900<=int(self.epoch.get())<=2100:
            messagebox.showerror(
            "Check input", "Allowed epoch from 1900 - 2100.")
            self.top.lift()
        elif TablesPDF().go(self.filename.get(), progress_bar, lat_min, lat_max, int(self.epoch.get())):
            self.top.destroy()

    def askpdffilename(self):
        self.filename.set('PUB249_Vol1_'+self.epoch.get()+'.pdf')
        self.filename.set(filedialog.asksaveasfilename(initialfile = self.filename.get(), defaultextension = '.pdf', filetypes = [("PDF Document",".pdf")]))
        if self.filename.get() == '':
            self.filename.set('PUB249_Vol1_'+self.epoch.get()+'.pdf')
        self.top.lift()

    def aboutwindow(self):
        messagebox.showinfo(
            "About this script", about)

    def makeentry(self, parent, caption, width=None, row=None, column=0, sticky=tk.E, rightof=None, columnspan=1, **options):
        if caption!='':
            label = tk.Label(parent, text=caption)
            label.grid(sticky=tk.E, row=row, column=column)
            row = label.grid_info()['row']
            column+=1
        elif rightof:
            row=rightof.grid_info()['row']
            column=int(rightof.grid_info()['column'])+1
        entry = tk.Entry(parent, justify=tk.RIGHT, **options)
        if width:
            entry.config(width=width)
        entry.grid(sticky=sticky,row=row, column=column, columnspan=columnspan)
        return entry
    
    def makeoutput(self, parent, caption, width=None, row=None, column=0, sticky=tk.E, rightof=None, returnlabel=False, columnspan=1, **options):
        textvar = tk.StringVar()
        if caption!='':
            label = tk.Label(parent, text=caption)
            label.grid(sticky=tk.E, row=row, column=column)
            row = label.grid_info()['row']
            column+=1
        elif rightof:
            row=rightof.grid_info()['row']
            column=int(rightof.grid_info()['column'])+1
        output = tk.Label(parent, textvar=textvar, anchor=tk.E, relief=tk.SUNKEN, **options)
        if width:
            output.config(width=width)
        output.grid(sticky=sticky,row=row, column=column, columnspan=columnspan)
        if returnlabel:
            return textvar, output
        else:
            return textvar
     
    def calculate(self):
        self.get()
        self.precomp = Precomputation(self.precData)
        self.precomp.compute()
        self.plot.draw(self.precomp)
        
        self.gha_out.set(str(self.precomp.gha))
        self.assumed_lon_out.set(str(self.precomp.assumed_lon))
        self.lha_out.set(str(self.precomp.lha_a))
        self.assumed_lat_out.set(str(self.precomp.assumed_lat))
        self.coriolis.set(str(round(self.precomp.coriolis, 2)))
        self.rhumbline.set(str(round(self.precomp.rhumbline, 2)))
        self.corrected_lat_out.set(str(self.precomp.corrected_lat))
        self.corrected_lon_out.set(str(self.precomp.corrected_lon))

        if (self.precData['body1']['name']=='' and
            self.precData['body2']['name']=='' and
            self.precData['body2']['name']==''):
            try:
                stars = GoodSet(self.precomp.observer).compute()
                if DEBUG.ON():
                    print('---- Automatic search - best solution ----')
                    for star in stars:
                        print(star.name, star.alt, star.az)
            except StopIteration as e:
                print(e)
            else:
                self.precData['body1']['name'] = stars[0].name
                self.precData['body2']['name'] = stars[1].name
                self.precData['body3']['name'] = stars[2].name               
                self.initialize()
        else:
            try:
                liste = [Stars(self.precData['body1'], self.precomp, self.precData['lower_limb']),
                         Stars(self.precData['body2'], self.precomp, self.precData['lower_limb']),
                         Stars(self.precData['body3'], self.precomp, self.precData['lower_limb'])]
            except Exception as e:
                # TODO: delete line ?
                raise e
                print(e)
            else:
                number = 0
                for i in liste:
                    self.body[number]['name'].set(str(i.name))
                    # only print the time portion and dicard the date
                    self.body[number]['date'].set(str(i.date).split()[1])
                    self.body[number]['gha'].set(str(i.gha))                    
                    self.body[number]['dec'].set(str(i.dec))
                    self.body[number]['zn'].set(str(i.zn))
                    self.body[number]['chc'].set(str(i.corr_hc))
                    self.body[number]['moo'].set(str(i.moo))
                    self.body[number]['mob'].set(str(i.mob))
                    self.body[number]['sd'].set(str(i.sd))
                    self.body[number]['pa'].set(str(i.pa))                    
                    self.body[number]['p_q'].set(str(i.q))
                    self.body[number]['refr'].set(str(i.refr))
                    self.body[number]['sext'].set(str(i.sextant_error))                    
                    
                    self.body[number]['tadj'].set(str(i.total_adj))
                    self.body[number]['ahc'].set(str(i.adj_hc))
                    self.body[number]['int'].set(str(round(i.intercept)))

                    # draw lines for Zn and intercept
                    if i.intercept == 0:
                        self.plot.intercept(number, self.precomp.corrected_lat, self.precomp.corrected_lon, i.zn)
                    else:
                        self.plot.intercept(number, self.precomp.corrected_lat, self.precomp.corrected_lon, i.zn, intercept = i.intercept)
                    
                    number+=1
                try:
                    intersections = find_int3(liste[0], liste[1], liste[2])
                except Exception as e:
                    self.output.config(text=('No intersection found, no Ho entered or you did not specify 3 bodies:'+str(e)))
                else:
                    self.output.config(text=(''))
                    self.plot.polygon(intersections)
                    lat, lon = least_error(intersections)
                    self.delta_out.set(str(round(180/pi*60*dist(lat, lon, self.precomp.dr_lat, self.precomp.dr_lon),2)))
                    self.mpp_lat_out.set(str(ephem.degrees(lat)))
                    self.mpp_lon_out.set(str(ephem.degrees(lon)))
                    # Dreieck mit Punkt fuer Fix zeichnen
                    self.plot.fix((lat, lon))

    def validate(self, action, index, value_if_allowed,
                       prior_value, text, validation_type, trigger_type, widget_name, datatype):
        limit = {}
        limit['year'] = {'lower':1900, 'upper':2100, 'empty':False}
        limit['year_epoch'] = {'lower':1900, 'upper':2100, 'empty':False}
        limit['year_pdf'] = {'lower':1900, 'upper':2100, 'empty':True}
        limit['month'] = {'lower':1, 'upper':12, 'empty':False}
        limit['day'] = {'lower':1, 'upper':31, 'empty':False}
        limit['hour'] = {'lower':0, 'upper':23, 'empty':False}
        limit['lat'] = {'lower':-89, 'upper':89, 'empty':False}
        limit['lon'] = {'lower':-179, 'upper':179, 'empty':False}
        limit['ho_dd'] = {'lower':0, 'upper':89, 'empty':True}
        limit['ho_mm'] = {'lower':0, 'upper':59, 'empty':True}
        limit['base60'] = {'lower':0, 'upper':59, 'empty':False}
        limit['base360'] = {'lower':0, 'upper':359, 'empty':False}
        limit['flgs'] = {'lower':0, 'upper':600, 'empty':False}
        limit['sext'] = {'lower':-20, 'upper':20, 'empty':False}
        if trigger_type=='key':
            self.invaliddata[datatype] = False
            # allow sign if lower limit < 0
            if limit[datatype]['lower']<0:
                characters = '0123456789+-'
            else:
                characters = '0123456789'
            # allow delete
            if int(action)==0:
                # allow empty field
                #if value_if_allowed=='':
                if len(value_if_allowed)<len(str(limit[datatype]['lower'])):
                    if not limit[datatype]['empty']:
                        self.invaliddata[datatype] = True
                return True
            # only allow specific characters specified above
            if text in characters or len(text)>1:
                # if first character is + or -
                if value_if_allowed[0]=='-' or value_if_allowed[0]=='+':
                    # only + or - ? Allow
                    if len(value_if_allowed)==1:
                        return True
                    sign = 1
                else:
                    sign = 0
                try:
                    int(value_if_allowed)
                except ValueError:
                    return False
                else:
                    # we need an additional check (or i.e. you can not enter a year if min is 1900)
                    if limit[datatype]['lower']>1 and len(value_if_allowed)<len(str(limit[datatype]['lower'])):
                        upper = int(str(limit[datatype]['upper'])[0:len(value_if_allowed)])
                        lower = int(str(limit[datatype]['lower'])[0:len(value_if_allowed)])
                        # even if value_if_allowed is ok so far as digits are missing set invaliddata
                        if not limit[datatype]['empty']:
                            self.invaliddata[datatype] = True
                        if upper>=int(value_if_allowed)>=lower:
                            return True
                        else:
                            return False
                    if limit[datatype]['upper']>=int(value_if_allowed)>=limit[datatype]['lower']:
                        return True
                    else:
                        if not limit[datatype]['empty']:
                            self.invaliddata[datatype] = True
                        return False
        return False

class GoodSet:
    """ seeks favorable constellation of bodies """
    def __init__(self, observer):
        self.observer = observer
        self.firstset = ()
        self.starlist = []
        for star in all_stars.values():
            star.compute(self.observer)
            if ephem.degrees('15')<star.alt<ephem.degrees('70'):
                # only use bodies that are above 10 and below 70 deg Hc
                self.starlist.append(star)


    def compute(self, number = 3):
        threestarfix = []
        if not number in [3,7]:
            raise Exception('Use with number = 3 or 7')
        if number == 7:
            # get three stars for a 3 star fix and remove them from the list of available stars
            self.firstset  = GoodSet(self.observer).compute()
            for star in self.firstset:
                threestarfix.append(star.name)
                self.starlist.remove(star)
        result = self.combinations(number = number)
        if len(result)==0:
            result = self.combinations(number = number, low = True)
        if len(result)==0:
            raise StopIteration('No objects found.')
        # sort by weight
        result.sort(key=lambda tup: tup[1])
        for stars, weight in result:#[0:15]:
            # sort by Zn
            result = sorted(stars, key=lambda star: star.az)
        if number == 7:
            result.extend(self.firstset)
            # return names of 3 best stars so they can be marked in the PDF
            return result, threestarfix
        return result

    def spacing(self, stars, low):
        min_angle = ephem.degrees('100')
        if low:
            min_angle = ephem.degrees('90.0')
        max_angle = ephem.degrees('150.0')
        dif1 = ephem.degrees(stars[2].az - stars[1].az)
        dif2 = ephem.degrees(stars[1].az - stars[0].az)
        dif3 = ephem.degrees(2*pi - (dif1 + dif2))
        if (min_angle<dif1<max_angle) and (min_angle<dif2<max_angle) and (min_angle<dif3<max_angle):
            return True
        return False

    def combinations(self, number = 3, low = False):
        result = []
        if number == 7:
            number_of_stars = 4
        else:
            number_of_stars = 3
        for star_combination in itertools.combinations(self.starlist, number_of_stars):
            my_stars = sorted(star_combination, key=lambda star: star.az)
            # Polaris is not used in Vol1 general tables
            if not 'Polaris' in [star.name for star in my_stars]:
                # check for even distribution in Zn
                if (self.spacing(my_stars, low) or number_of_stars == 4):
                    # weight is the sum of magnitudes, the smaller the better
                    weight = 0
                    for star in my_stars:
                                           weight -= star.mag
                    result.append((my_stars, weight))
        return result


def main(args):
    global all_stars
    all_stars = {}
    for line in db.strip().split('\n'):
        star = ephem.readdb(line)
        all_stars[star.name] = star

    
    # used for debugging PDF only
    #TODO fix epoch - what exactly is epoch.0
    """
    try:
        report = TablesPDF()
    except:
        pass
    else:
        report.go('test.pdf', None, 10, 12, ephem.date('2010/1/1'))
    """
    root = tk.Tk()
    gui = Gui(master = root)
    gui.mainloop()
    
    # TODO: calculate PREC/NUT??? difference between epoch as vector?
    return 0

if __name__ == '__main__':
    try:
        sys.exit(main(sys.argv))
    except Exception as e:
        messagebox.showerror(
            "Fatal error",
            "Exception occurred:\n%s" % e)
    
