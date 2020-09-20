import numpy as np
import matplotlib.pyplot as plt
import pickle
from scipy import ndimage
from datetime import datetime
from matplotlib.patches import Rectangle # only used for plotting in setTRLed
import clr, os, winreg
from itertools import islice
import pdb





#
#%%
""" TODO
- Physical optics. Set settings. Run analysis. Save data/image
- Merit function weight adjust. Adjust weight for part of the MF


"""


#%%

def zInitInteractive():

    """
    Function for initiation

    Parameters: -
    Returns: (TheConnection, TheApplication, TheSystem, ZOSAPI)
    Notes:
    Examples:
        (TheConnection, TheApplication, TheSystem, ZOSAPI) = zInitInteractive()
        TheSystemData=TheSystem.SystemData
    """

    aKey = winreg.OpenKey(winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER), r"Software\Zemax", 0, winreg.KEY_READ)
    zemaxData = winreg.QueryValueEx(aKey, 'ZemaxRoot')
    NetHelper = os.path.join(os.sep, zemaxData[0], r'ZOS-API\Libraries\ZOSAPI_NetHelper.dll')
    winreg.CloseKey(aKey)
    clr.AddReference(NetHelper)
    import ZOSAPI_NetHelper
    print(ZOSAPI_NetHelper)

    pathToInstall = ''
    # uncomment the following line to use a specific instance of the ZOS-API assemblies
    #pathToInstall = r'C:\C:\Program Files\Zemax OpticStudio'

    success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(pathToInstall);
    zemaxDir = ''
    if success:
        zemaxDir = ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory();
        # print('Found OpticStudio at:   %s' + zemaxDir);
    else:
        raise Exception('Cannot find OpticStudio')
    clr.AddReference(os.path.join(os.sep, zemaxDir, r'ZOSAPI.dll'))
    clr.AddReference(os.path.join(os.sep, zemaxDir, r'ZOSAPI_Interfaces.dll'))
    import ZOSAPI
    TheConnection = ZOSAPI.ZOSAPI_Connection()
    if TheConnection is None:
        raise Exception("Unable to intialize NET connection to ZOSAPI")
    TheApplication = TheConnection.ConnectAsExtension(0)
    if TheApplication is None:
        raise Exception("Unable to acquire ZOSAPI application")
    if TheApplication.IsValidLicenseForAPI == False:
        raise Exception("License is not valid for ZOSAPI use.  Make sure you have enabled 'Programming > Interactive Extension' from the OpticStudio GUI.")
    TheSystem = TheApplication.PrimarySystem
    if TheSystem is None:
        raise Exception("Unable to acquire Primary system")

    print('SystemID:', TheSystem.SystemID,
      ' Mode:', TheSystem.Mode,
      '',TheSystem.SystemFile, '\n'
      )

    return TheConnection, TheApplication, TheSystem, ZOSAPI

def reshape(self, data, x, y, transpose = False):
    """Converts a System.Double[,] to a 2D list for plotting or post processing

    Parameters
    ----------
    data      : System.Double[,] data directly from ZOS-API
    x         : x width of new 2D list [use var.GetLength(0) for dimension]
    y         : y width of new 2D list [use var.GetLength(1) for dimension]
    transpose : transposes data; needed for some multi-dimensional line series data

    Returns
    -------
    res       : 2D list; can be directly used with Matplotlib or converted to
                a numpy array using numpy.asarray(res)
    """
    if type(data) is not list:
        data = list(data)
    var_lst = [y] * x;
    it = iter(data)
    res = [list(islice(it, i)) for i in var_lst]
    if transpose:
        return self.transpose(res);
    return res

def transpose(data):
    """Transposes a 2D list (Python3.x or greater).

    Useful for converting mutli-dimensional line series (i.e. FFT PSF)

    Parameters
    ----------
    data      : Python native list (if using System.Data[,] object reshape first)

    Returns
    -------
    res       : transposed 2D list
    """
    if type(data) is not list:
        data = list(data)
    return list(map(list, zip(*data)))

# EX: (surfNum [int], surf [ILDERow]) = zGetSEQSurf(TheSystem, 'surfaceName' ,printSurfacesFound=True)
def zGetSEQSurf(TS, comment_, printSurfacesFound=False):



    LDE=TS.LDE
    for ss in range(0,LDE.NumberOfSurfaces):
        # print(LDE.GetSurfaceAt(ss).Comment)
        if LDE.GetSurfaceAt(ss).Comment==comment_:
            if printSurfacesFound: print('Found: ', ss, '', comment_, '')
            return (ss, LDE.GetSurfaceAt(ss)) # Return the surf number and object
    print('Error: ', comment_, ' not found')
    return False

# (objNum [int], obj [INCERow]) = zGetNSCObject(TheSystem, 'objName' , printSurfacesFound=True)
def zGetNSCObject(TS, comment_, printSurfacesFound=True):
    TheNCE=TS.NCE
    for nn in range(0,TheNCE.NumberOfObjects+1):
        oo = TheNCE.GetObjectAt(nn)
        if oo.Comment==comment_:
            print('Found: ', nn, '', comment_, '')
            return (nn, oo)
    print('Error: ', comment_ ,' not found')
    return False

# Ex: zNSCRaytrace(TheSystem, splitRays=False, scatterRays=False, usePolarization=False, ignoreErrors=False, saveRays='JCRAYSFilt01.ZRD', rayFilter='H2', clearDetectors=True, numberOfCores=8)
#saveRays: False / 'filename'
def zNSCRaytrace(TS, splitRays=False, scatterRays=False, usePolarization=False, ignoreErrors=False, saveRays=False, saveRaysFilename='rays01.ZRD', rayFilter='', clearDetectors=True, numberOfCores=8, printStatus=True):

    """
    Run a Non-sequential raytracing

    Parameters: (TheSystem, optional_varaibles)
    Returns: True
    Notes:
    Examples:
        zNSCRaytrace(TheSystem)
    ToDo:
        Random seed
        Beep sound when ready
        Return something that makes sense. Detector energy?
    """

    if splitRays: usePolarization=True
    NSCRayTrace = TS.Tools.OpenNSCRayTrace()
    NSCRayTrace.NumberOfCores=numberOfCores
    NSCRayTrace.UsePolarization = usePolarization
    NSCRayTrace.IgnoreErrors = ignoreErrors
    NSCRayTrace.SplitNSCRays = splitRays
    NSCRayTrace.ScatterNSCRays = scatterRays
    if saveRays:
        NSCRayTrace.SaveRays = True
        NSCRayTrace.SaveRaysFile = saveRays
        NSCRayTrace.Filter=rayFilter
        sr='Saving rays: '+str(saveRays)
    else:
        sr=''
        NSCRayTrace.SaveRays = False

    if clearDetectors: NSCRayTrace.ClearDetectors(0)
    NSCRayTrace.Run()
    lastValue = []
    lastValue.append(0)
    while NSCRayTrace.IsRunning:
        currentValue = NSCRayTrace.Progress
        if currentValue % 2 == 0:
            if lastValue[len(lastValue) - 1] != currentValue:
                lastValue.append(currentValue)
                print('\b\b\b\b\b\b\b\b\b\b\b\b\b\b\b\b\b\bRaytracing: ', currentValue, '%', end="")

    NSCRayTrace.WaitForCompletion()
    print('\nDead rays', NSCRayTrace.DeadRayErrors, '  DeadRayTreshold: ', NSCRayTrace.DeadRayThreshold)
    print(sr)
    NSCRayTrace.Close()

    return True


def zObjAny2ObjNum(TS,objAny):
    """

    Parameters: TheSystem, object_comment_string / object_number / NSC_object
    Returns: NSC object number
    Notes:
    Examples:
        zObjAny2ObjNum(TheSystem, 'myDetector') → 8
    ToDo:
    """

    if type(objAny)==str: # detObj can be a comment, row number, or object
        (objNo, objAny)=zGetNSCObject(TS, objAny)
    elif type(obj)==int:
        objNo=objAny
    else:
        objNo=objAny.ObjectNumber
    return objNo


def zObjAny2ObjObj(TS,objAny):
    """

    Parameters: TheSystem, object_comment_string / object_number / NSC_object
    Returns: NSC object
    Notes:
    Examples:
        zObjAny2ObjNum(TheSystem, 'myDetector') → NSC_object
    ToDo:
    """
    if type(objAny)==str: # detObj can be a comment, row number, or object
        (objNo, obj)=zGetNSCObject(TS, objAny)
    elif type(obj)==int:
        (objNo, obj)=zGetNSCObject(TS, objAny)
        TheNCE=TS.NCE
        obj=TheNCE.GetObjectAt(objAny)
    else:
        obj=obj.ObjectNumber
    return obj

def zReadDetector(TS, detObj, smoothing=0, plotData=True):
    """

    Parameters:
        TheSystem
        detobject_comment_string / object_number / NSC_object
        ..
    Returns: detector data
    Notes:
    Examples:
        zReadDetector(TheSystem, 'myDetector') → 88
    ToDo:
        Plot contour lines
    """

    print('Getting detector data')
    objNo=zObjAny2ObjNum(TS,detObj)
    obj = TS.NCE.GetObjectAt(objNo)
    TheNCE=TS.NCE
    numXPixels = obj.ObjectData.NumberXPixels
    numYPixels = obj.ObjectData.NumberYPixels
    pltWidth   = 2 * obj.ObjectData.XHalfWidth
    pltHeight  = 2 * obj.ObjectData.YHalfWidth
    pix = 0
    detectorData = [[0 for x in range(numYPixels)] for x in range(numXPixels)]
    for x in range(0,numYPixels,1):
        for y in range(0,numXPixels,1):
            ret, pixel_val = TheNCE.GetDetectorData(objNo, pix, 1, 0)
            pix += 1
            if ret == 1:
                detectorData[y][x] = pixel_val
            else:
                detectorData[x][y] = -1
    # detectorData = transpose(detectorData)
    detectorData=np.array(detectorData)
    detectorData=np.rot90(detectorData)

    if smoothing>0: detectorData = ndimage.uniform_filter(detectorData, size=smoothing)

    plotData=True
    if plotData:
        fig=plt.figure(1)
        fig.clf()
        ax = fig.add_subplot(111)
        h1=ax.imshow(detectorData, cmap='jet', aspect='equal', extent=(-pltWidth*0.5, pltWidth*0.5, -pltHeight*0.5, pltHeight*0.5))
        ax.set_xlabel('[mm]', fontsize=8)
        ax.set_ylabel('[mm]', fontsize=8)
    return detectorData



def zSetNSCParameter(TS, objAny, par, val):
    """

    Parameters:
        TheSystem
        NSC_object
        parameter
        value
        ..
    Returns: -
    Notes:
    Examples:
        zSetNSCParameter(TheSystem, 'myLens', 3, 1)
    ToDo:
        support for zpos, material mm

    """

    TheNCE=TS.NCE
    objNo=zObjAny2ObjNum(TS, objAny)
    objCell=TheNCE.GetObjectAt(objNo).GetObjectCell(par+10)
    dataType=objCell.DataType # 0=int; 1=double; 2=string
    if dataType==0:
        objCell.IntegerValue=val
    elif dataType==1:
        objCell.DoubleValue=val
    else:
        objCell.Value=val


# Start At: -1: DMFS; 0: Last line
def zSEQOptWizard(TS, Criterion='Spot', SpatialFrequency=100, XSWeight=1, YTWeigh=1, Type='RMS', Reference='Centroid', UseGaussianQuadrature=True, UseRectangularArray=False, Rings=3, Arms='6arms', Obscuration=0, GridSizeNxN=4, DeleteVignetted=True, UseGlassBoundaryValues=False, GlassMin=0, GlassMax=1000, GlassEdgeThickness =0, UseAirBoundaryValues=False, AirMin=0, AirMax=1000, AirEdgeThickness =0, StartAt=1, OverallWeight=1, ConfigurationNumber=1, UseAllConfigurations =False, FieldNumber=1, UseAllFields=False, AssumeAxialSymmetry=False, IgnoreLateralColor=False, AddFavoriteOperands=False, OptimizeForManufacturingYield=False, ManufacturingYieldWeight=0, MaxDistortionPct=False):


"""
    Generate Merit Function using Wizard

    Parameters:
        TheSystem
        Criterion: 'Wavefront', 'Contrast', 'Spot', 'Angular'
        ..
        Type: 'RMS', 'PTV'
        Reference: 'Centroid', 'ChiefRay', 'Unreferenced'
        ..
        Arms: '6arms', '8arms', '10arms', '12arms'
        ..
        StartAt: -1 (starts at first DMFS row); 0 (stars at last line in MF); other (starts at that line)
    Returns: -
    Notes:
    Examples:
            zSEQOptWizard(TheSystem, ConfigurationNumber=2, OverallWeight=1, StartAt=-1, Criterion='Spot', AssumeAxialSymmetry=False, Rings=5)
    ToDo:
        Add BLNK Tags with '* IMG_CONF#..' just after the configuration change, so that the MF weight check  automatically can tag this region
        How to do distinction between wiz1 and wiz2?
        Ignore lateral color option is not available in the interface → email zemax support

    """


    TheMFE = TS.MFE
    OptWizard = TheMFE.SEQOptimizationWizard2

    # Criterion
    criterionDict={'Wavefront':0, 'Contrast':1, 'Spot':2, 'Angular':3}
    crit=criterionDict.get(Criterion, 'nan')
    OptWizard.Criterion=crit  # 0 Wavefront; 1 Contrast; 2 Spot; 3 Angular

    # Spatial frequency for Contrast optimization
    if crit==1:
        OptWizard.SpatialFrequency=SpatialFrequency
        OptWizard.XSWeight=XSWeight
        OptWizard.YTWeigh=YTWeigh

    # Optmization Type
    typeDict={'RMS':0, 'PTV':1}
    typ=typeDict.get(Type, 'nan')
    OptWizard.Type=typ # 0 RMS; 1 PTV

    # Reference
    referenceDict={'Centroid':0, 'ChiefRay':1, 'Unreferenced':2}
    ref=referenceDict.get(Reference, 'nan')
    OptWizard.Reference=ref # 0 Centroid; 1 ChiefRay; 2 Unreferenced

    # Pupil Integration
    OptWizard.UseGaussianQuadrature=UseGaussianQuadrature
    OptWizard.UseRectangularArray=UseRectangularArray
    OptWizard.Rings=Rings
    armsDict={'6arms':0, '8arms':1, '10arms':2, '12arms':3}
    arms=armsDict.get(Arms, 'nan')
    OptWizard.Arms=arms # 0 6arms; 1 8 arms; 2 10arms; 3 12 arms
    OptWizard.Obscuration=0
    OptWizard.GridSizeNxN=4 # 4...204
    OptWizard.DeleteVignetted=True

    # Boundary values
    OptWizard.UseGlassBoundaryValues=UseGlassBoundaryValues
    OptWizard.GlassMin=GlassMin
    OptWizard.GlassMax=GlassMax
    OptWizard.GlassEdgeThickness =GlassEdgeThickness
    OptWizard.UseAirBoundaryValues=UseAirBoundaryValues
    OptWizard.AirMin=AirMin
    OptWizard.AirMax=AirMax
    OptWizard.AirEdgeThickness =AirEdgeThickness

    # Find MF row with first DMSF
    if StartAt==-1:
        StartAt=0
        foundDMFS=False
        for rowNo in np.arange(1, TheMFE.NumberOfOperands+1):
            Operand=TheMFE.GetOperandAt(rowNo)
            # contrib=Operand.Contribution

            if Operand.TypeName=='DMFS':
                foundDMFS=True
                StartAt=rowNo
                print('Found DMFS at row ', StartAt)
                break
        if not foundDMFS: print('No DMFS found. Starting MF at last row.')
    OptWizard.StartAt=StartAt

    #
    OptWizard.OverallWeight=OverallWeight
    OptWizard.ConfigurationNumber=ConfigurationNumber
    OptWizard.UseAllConfigurations =UseAllConfigurations
    OptWizard.FieldNumber=FieldNumber
    OptWizard.UseAllFields=UseAllFields
    OptWizard.AssumeAxialSymmetry=AssumeAxialSymmetry
    OptWizard.IgnoreLateralColor=IgnoreLateralColor
    OptWizard.AddFavoriteOperands=AddFavoriteOperands

    # Manufacturing Yield
    if OptimizeForManufacturingYield:
        OptWizard.OptimizeForBestNominalPerformance=False
        OptWizard.OptimizeForManufacturingYield=True
        OptWizard.ManufacturingYieldWeight=ManufacturingYieldWeight
    else:
        OptWizard.OptimizeForManufacturingYield=False
        OptWizard.OptimizeForBestNominalPerformance=True

    # Max Distortion
    if MaxDistortionPct:
        OptWizard.UseMaximumDistortion=True
        OptWizard.MaxDistortionPct=MaxDistortionPct
    else:
        OptWizard.UseMaximumDistortion=False
    OptWizard.OK()



def zSEQMakeVariable(TS, surfList, variableOrFixed='variable', radiusOrThic='radius'):
    """
    Change LDE surface radius or thicknes to variable or fixed
    Parameters:
        TheSystem
        List of surfaces, for example: (1,3,5)
        variableOrFixed: 'variable', 'fixed'
        radiusOrThic: 'radius', 'thic'

    Returns: -
    Notes:
    Examples:
        zSEQMakeVariable(TheSystem, (1,4,5), variableOrFixed='variable')
    ToDo:
    """

    columnDict={'radius':2, 'thic':3}
    col=columnDict.get(radiusOrThic, 'nan')

    for ss in surfList:
        (surfNum, surf)=zGetSEQSurf(TS, ss)
        # print(zGetSEQSurf(TS, ss))
        # print(surf.GetSurfaceCell(2).MakeSolveVariable())
        if variableOrFixed=='variable':
            surf.GetSurfaceCell(col).MakeSolveVariable()
        elif variableOrFixed=='fixed' or variableOrFixed=='f' :
            surf.GetSurfaceCell(col).MakeSolveFixed()
        else:
            print('Error. Pls specify variable or fixed')

def zSystemSettingts(TS, settingsNo):


    """
    Change system settings, for example wavelengths, aperture, ray aiming

    Parameters:
        TheSystem
        settingsNo: [int] representing a collection of system settings

    Returns: -
    Notes:
    Examples:
        zSEQMakeVariable(TheSystem, (1,4,5), variableOrFixed='variable')
    ToDo:
    """
    SystExplorer = TS.SystemData

    if settingsNo==1:
        SystExplorer.Aperture.ApertureType=0 # 0:ENPD ; 1=ImageSpaceFNO ; ..
        SystExplorer.Aperture.ApertureValue = 45
        SystExplorer.RayAiming.RayAiming=ZOSAPI.SystemData.RayAimingMethod.Off #Off / Paraxial / Real
        SystExplorer.TitleNotes.Title='TitleTest'
        SystExplorer.TitleNotes.Notes='Some notes'

        # Fields
        SystExplorer.Fields.DeleteAllFields()
        SystExplorer.Fields.AddField(0, 0.5, 0.5)

        # Wavelengths
        SystExplorer.Wavelengths.GetWavelength(1).Wavelength = 0.55
        SystExplorer.Wavelengths.SelectWavelengthPreset(ZOSAPI.SystemData.WavelengthPreset.FdC_Visible )
        SystExplorer.Wavelengths.AddWavelength(0.51, 0.5)
        SystExplorer.Wavelengths.AddWavelength(0.555, 1)
        SystExplorer.Wavelengths.AddWavelength(0.61, 0.5)
        SystExplorer.Wavelengths.RemoveWavelength(1)
        SystExplorer.Wavelengths.RemoveWavelength(1)
        SystExplorer.Wavelengths.RemoveWavelength(1)
        print('settingsNo=1')
        return True

    elif settingsNo==2:
        print('TBD')
    else:
        print("Error: No such setting number")
        return False



def zMFClocker(TS, samplingTime=5, algorithm='hammer'):
    """
    Use for testing the calculation time for a Merit Function


    Parameters:
        TheSystem
        samplingTime: how long to run the calculation for

    Returns: milliseconds calculation time per system
    Notes:
    Examples:
        zMFClocker(TheSystem)
    ToDo:
        Add support for normal optimization (not hammer)
    """

    t0 = datetime.now().timestamp()
    HammerOpt = TS.Tools.OpenHammerOptimization()
    HammerOpt.RunAndWaitWithTimeout(samplingTime)
    HammerOpt.Cancel()
    HammerOpt.WaitForCompletion()
    noOfSyst=HammerOpt.Systems
    msPerSyst=samplingTime*1000/noOfSyst
    print(f'Average time per system: {msPerSyst:.1f} ms. Number of systems {noOfSyst:.0f}  ')
    HammerOpt.Close()

    return msPerSyst

def zOptimize(TS, optimizationType='Hammer', timeStepMinutes=20, totalRuntimeMinutes=10*60, numberOfCores=8, saveSystem=True):

    """
    Run optmization


    Parameters:
        TheSystem
        optimizationType: currently supports only 'Hammer'
        timeStepMinutes: Time to run the optimzation before reporting MF value
        totalRunTimeMinutes: Time to run the optimization before stop
        numberOfCores:
        saveSystem:

    Examples:
        zOptimize(TheSystem, timeStepMinutes=10, totalRuntimeMinutes=100)
    ToDo:
        How to break execution in a nice way?
        Support for global optimization
        Plot MF % live, logscale
        Send notifications to mobile phone of current MF value

    """

    dt = datetime.now().timestamp()
    print('Running Hammer Optimization')
    TheMFE = TS.MFE
    HammerOpt = TS.Tools.OpenHammerOptimization()
    HammerOpt.NumberOfCores=8
    mfList=[]
    mmList=[]
    mfStarval=TheMFE.CalculateMeritFunction()
    print('MFValueStart: ', mfStarval)
    mfLastValue=mfStarval
    dtElapsed=datetime.now().timestamp()-dt
    print('MFvalue\t\tMFRelative   Timestamp\tImprovement')
    steps=np.round(totalRuntimeMinutes/timeStepMinutes)

    try:
        for ss in np.arange(0,steps):
                improvementPct=0
                improvementStr=''
                HammerOpt.RunAndWaitWithTimeout(timeStepMinutes*60)
                HammerOpt.Cancel()
                HammerOpt.WaitForCompletion()
                mf=TheMFE.CalculateMeritFunction()
                mfPct=mf/mfStarval
                # if mf<mfLastValue: improvement='*******'
                improvementPct=1*(1-(mf/mfLastValue))
                if improvementPct > 0 : improvementStr='   ******'
                dtElapsed=datetime.now().timestamp()-dt
                print( f'{mfLastValue:.6e}\t{mfPct:.9f}\t{dtElapsed/60:.1f}\t{improvementPct:.9f}\t'+improvementStr)

                mfList.append(mf)
                mmList.append(ss)
                mfLastValue=mf
    #             ax.plot(mmList, mfList, '-o')
    #             # plt.show()
    #             fig.canvas.draw()
                if saveSystem: TS.Save()
        HammerOpt.Close()

    except:
        HammerOpt.Close()
        print('Breaking Execution')


# In MFE use * in comment to make region eg.: * TRACKER
def zMFContributions(TS, plotFigure=True, plotLogscale=False):

    """
    Calculates the summed MF value for different sections in the MFE. Sections are marked with text starting with * . For example:
    BLNK  * LensDimensions
    CTGT  ...

    Parameters:
        TheSystem

    Examples:
        zMFContributions(TheSystem)
    ToDo:
        handle cases when BLNK is blank
        Add top list of which lines are worst contributors, color by type,
        100 first, bars to show how large they are, show in which conf they are,

    """

    TheMFE = TS.MFE
    contribSum=0
    contribSumList=[]
    contribTagList=[]
    whichConfAmIIn=''
    TheMFE.CalculateMeritFunction()
    confNo=''
    for rowNo in np.arange(1, TheMFE.NumberOfOperands+1):
        Operand=TheMFE.GetOperandAt(rowNo)
        contrib=Operand.Contribution
        if np.isfinite(contrib):
            contribSum=contribSum+contrib

        if Operand.TypeName=='CONF':
            confNo=Operand.GetOperandCell(2).IntegerValue

        if Operand.TypeName=='BLNK':
            tag=Operand.GetOperandCell(1).Value
            if len(tag)>0:
                if tag[0]=='*':
                    contribSumList.append(contribSum)
                    contribTagList.append('conf '+str(confNo)+'\n'+ tag[1:] )
                    contribSum=0

        if rowNo==TheMFE.NumberOfOperands:
            contribSumList.append(contribSum)

# print(rowNo, '\t', TheMFE.GetOperandAt(rowNo).Contribution,'\t',  contribSum, '\t', tag)
# print(contribSumList, '\n', contribTagList)

    contribSumList=contribSumList[1:]  # Remove first element in list
    contribSumList.reverse()
    contribTagList.reverse()

    for ii, cl in enumerate(contribSumList):  # Add percentage to tag names
        contribTagList[ii]=contribTagList[ii]+ f'\n {cl:.1f}%'

    if plotFigure:
        fig=plt.figure(1)
        fig.clf()
        ax = fig.add_subplot(111)
        lpx= np.arange(len(contribSumList))   # label position x
        bw=0.25  # bar width
        rects=ax.barh(lpx-bw, contribSumList, label='PeakIrradiance', tick_label=contribTagList, color='cadetblue')
        ax.set_title('Merit Function Contributions', fontsize=18)
        for tick in ax.xaxis.get_major_ticks(): tick.label.set_fontsize(13)
        for tick in ax.yaxis.get_major_ticks(): tick.label.set_fontsize(13)

        if plotLogscale:  ax.set_yscale('log')

    print(contribSumList)


# zHideNSCObjects(TheSystem, rStart, rStop)
def zHideNSCObjects(TS, objStart, objStop):
    """
    Hides NSC objects

    Parameters:
        TheSystem
        objStart: row number of first object to hide
        objStop:

    Examples:
        zHideNSCObjects(TheSystem, 6,8)
    ToDo:

    """

    TheMCE=TS.MCE
    for r in np.arange(objStop, objStart-1,-1):
        TheMCE.InsertNewOperandAt(1)
        TheMCE.InsertNewOperandAt(2)
        MCOperand1 = TheMCE.GetOperandAt(1)
        MCOperand2 = TheMCE.GetOperandAt(2)
        MCOperand1.ChangeType(ZOSAPI.Editors.MCE.MultiConfigOperandType.NPRO)
        MCOperand2.ChangeType(ZOSAPI.Editors.MCE.MultiConfigOperandType.NPRO)
        MCOperand1.Param2 = r
        MCOperand1.Param3 = 3
        MCOperand2.Param2 = r
        MCOperand2.Param3 = 4
        MCOperand1.GetOperandCell(1).IntegerValue=1
        MCOperand2.GetOperandCell(1).IntegerValue=1
        print('Hiding surf ', r)


def zChangeMFWeights(TS, rowStart, rowStop, weight):
    TheMFE = TS.MFE
    for rowNo in np.arange(rowStart, rowStop+1):
        Operand=TheMFE.GetOperandAt(rowNo)
        Operand.Weight=weight

def zChangeMFTargets(TS, rowStart, rowStop, target):
    TheMFE = TS.MFE
    for rowNo in np.arange(rowStart, rowStop+1):
        Operand=TheMFE.GetOperandAt(rowNo)
        Operand.Target=target



#%% Useful code
"""

def meritfcn...:
    _EFFL=ZOSAPI.Editors.MFE.MeritOperandType.EFFL
    TheSystem.MFE.GetOperandValue(_EFFL, 0, 2, 0,0,0,0,0,0)


    """
