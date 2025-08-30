import json
import math
from sympy import Symbol,nsolve,exp
import numpy as np
import re

def decomp(phi_0, c, top_p, bottom_p, top_decomp):
    """Performs decompaction based on the following parameters:
    (Initial porosity, Coefficient c, Present top depth, Present bottom depth, Decompacted top depth)"""
    thickness_p = bottom_p - top_p
    if thickness_p == 0:
        return 0
    center_p = (bottom_p + top_p)/2
    phi_p = phi_0 * math.exp(-center_p*c)
    d = Symbol('d')
    f = (1-phi_p) * thickness_p / (1-(phi_0 * exp(-(d + top_decomp)*c))) + (2*d)
    return float(nsolve(f,d,0))

'''
(Invoke-WebRequest -Uri https://bootstrap.pypa.io/get-pip.py -UseBasicParsing).Content | .\python.exe -
.\python.exe -m pip install sympy
.\python.exe -m pip install numpy


Function decomp(phi_0, c, top_p, bottom_p, top_decomp As Single) As Single 
    pyFile = "BasinVisPy-LibreOffice.py"  
    pyFunc = "decomp"  
    pyParams = Array(phi_0, c, top_p, bottom_p, top_decomp)  
      
    ScriptProvider = CreateUNOService("com.sun.star.script.provider.MasterScriptProviderFactory").createScriptProvider("")  
    pyScript = ScriptProvider.getScript("vnd.sun.star.script:" & pyFile & "$" & pyFunc & "?language=Python&location=user")  
    decomp = pyScript.invoke(pyParams, Array(), Array())  
End Function  

'''
