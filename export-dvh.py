import sys
import csv
import clr
import System

# Add Varian's path when referencing assemblies
sys.path.append(r"C:\Program Files (x86)\Varian\Vision\13.6\Bin64")

# Reference the Varian API assemblies
clr.AddReferenceToFile("VMS.TPS.Common.Model.API.dll")
clr.AddReferenceToFile("VMS.TPS.Common.Model.Types.dll")

from VMS.TPS.Common.Model.API import *
from VMS.TPS.Common.Model.Types import *

# Check that the correct number of arguments were given
# Note: There is always at least one argument because
#       sys.argv[0] is the script file name
if len(sys.argv) != 3:
    print 'Arguments: patient-id output-file-name'
    exit(1)

patientId = sys.argv[1]
outputFileName = sys.argv[2]

# Log in to Eclipse (shows the log-in dialog box)
app = Application.CreateApplication(None, None)

patient = app.OpenPatientById(patientId)

# Open the output file for writing
with open(outputFileName, 'w') as outputFile:

    # Create the object used to write in CSV format
    dvhWriter = csv.writer(outputFile)

    # Go through every Course, PlanSetup, and Structure
    # TODO: Handle cases where any of these is None
    for course in patient.Courses:
        for plan in course.PlanSetups:
            for structure in plan.StructureSet.Structures:

                # Get the cumulative DVH in absolute units
                dvhData = plan.GetDVHCumulativeData(structure,
                    DoseValuePresentation.Absolute,
                    VolumePresentation.AbsoluteCm3, 0.01)

                # dvhData is None if the DVH could not be calculated
                if dvhData != None:
                    # Write every DVH point in CSV format
                    for dvhPoint in dvhData.CurveData:
                        dvhWriter.writerow([patientId, plan.Id, structure.Id, \
                            dvhPoint.DoseValue.Dose, dvhPoint.Volume])
                else:
                    # Print to the error stream (normally the screen)
                    print >> sys.stderr, 'Cannot get DVH for structure', \
                        structure.Id, 'in plan', plan.Id

# Free unmanaged resources
app.Dispose()
