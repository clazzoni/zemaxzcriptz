# zosPython

zosPython.py is a collection of functions that can be used for interacting with Zemax OpticStudio

NB: Code is just for inspiration for custom scripts, etc. Lots of work left to do..

## Installation

Start by installing zopapi and pythonnet:

```bash
pip install zosapi
python -m pip install pythonnet
```
[START] > environmental variables >


Examples of usage:

```bash

import numpy as np
import matplotlib.pyplot as plt
from zosPython import *


(TheConnection, TheApplication, TheSystem, ZOSAPI) = zInitInteractive()
TheSystemData=TheSystem.SystemData


TheSystemData=TheSystem.SystemData
zGetSEQSurf(TheSystem, 'dd',printSurfacesFound=True)
zGetNSCObject(TheSystem, 'Field 1', printSurfacesFound=True)
zNSCRaytrace(TheSystem, splitRays=False, scatterRays=False, usePolarization=False, ignoreErrors=False, saveRays='JCRAYSFilt01.ZRD', rayFilter='H2', clearDetectors=True, numberOfCores=8)
zNSCRaytrace(TheSystem)
(objNum , obj ) = zGetNSCObject(TheSystem, 'dd')
dd=zReadDetector(TheSystem, obj)

zSetNSCParameter(TheSystem, 'ff', 1 , 12.34)
TheNCE=TheSystem.NCE
oo=TheNCE.GetObjectAt(4)


zSEQOptWizard(TheSystem, Criterion='Wavefront', SpatialFrequency=5123, Type='RMS' , Reference='Centroid',
              UseGaussianQuadrature='True', Rings=6, Arms='8arms'5, StartAt=27,  OptimizeForManufacturingYield=True, ManufacturingYieldWeight=0.5, MaxDistortionPct=False)

zSEQMakeVariable(TheSystem, ['aa', 'ee'], variableOrFixed='variable', radiusOrThic='thic')
zSEQMakeVariable(TheSystem, ['aa', 'cc'], variableOrFixed='fixed', radiusOrThic='radius')

zSystemSettingts(TheSystem, 24)
et=zMFClocker(TheSystem, noOfMFCalls=1)

zOptimize(TheSystem, timeStepMinutes=3/60, totalRuntimeMinutes=12/60)

zMFContributions(TheSystem)
```
