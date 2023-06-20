import click
from pcbnewTransition import pcbnew
import csv
import os
import re
import sys
import shutil
from datetime import datetime as dt
from pathlib import Path
from kikit.fab.common import *
from kikit.common import *
from kikit.export import gerberImpl, exportSettingsPcbway

def collectSolderTypes(board):
    result = {}
    for footprint in board.GetFootprints():
        if excludeFromPos(footprint):
            continue
        if hasNonSMDPins(footprint):
            result[footprint.GetReference()] = "thru-hole"
        else:
            result[footprint.GetReference()] = "SMD"

    return result

def addVirtualToRefsToIgnore(refsToIgnore, board):
    for footprint in board.GetFootprints():
        if excludeFromPos(footprint):
            refsToIgnore.append(footprint.GetReference())

def collectBom(components, manufacturerFields, partNumberFields,
               descriptionFields, notesFields, typeFields, footprintFields,
               ignore):
    bom = {}

    # Use KiCad footprint as fallback for footprint
    footprintFields.append("Footprint")
    # Use value as fallback for description
    descriptionFields.append("Value")

    for c in components:
        if getUnit(c) != 1:
            continue
        reference = getReference(c)
        if reference.startswith("#PWR") or reference.startswith("#FL") or reference in ignore:
            continue
        if hasattr(c, "in_bom") and not c.in_bom:
            continue
        if hasattr(c, "on_board") and not c.on_board:
            continue
        if hasattr(c, "dnp") and c.dnp:
            continue
        manufacturer = None
        for manufacturerName in manufacturerFields:
            manufacturer = getField(c, manufacturerName)
            if manufacturer is not None:
                break
        partNumber = None
        for partNumberName in partNumberFields:
            partNumber = getField(c, partNumberName)
            if partNumber is not None:
                break
        description = None
        for descriptionName in descriptionFields:
            description = getField(c, descriptionName)
            if description is not None:
                break
        notes = None
        for notesName in notesFields:
            notes = getField(c, notesName)
            if notes is not None:
                break
        solderType = None
        for typeName in typeFields:
            solderType = getField(c, typeName)
            if solderType is not None:
                break
        footprint = None
        for footprintName in footprintFields:
            footprint = getField(c, footprintName)
            if footprint is not None:
                break

        cType = (
            description,
            footprint,
            manufacturer,
            partNumber,
            notes,
            solderType
        )
        bom[cType] = bom.get(cType, []) + [reference]
    return bom

def bomToCsv(bomData, filename, nBoards, types):
    with open(filename, "w", newline="", encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["Item #", "Designator", "Qty", "Manufacturer",
                         "Mfg Part #", "Description / Value", "Footprint",
                         "Type", "Your Instructions / Notes"])
        item_no = 1

        tmp = {}
        for cType, references in bomData.items():
            tmp[references[0]] = (references, cType)

        for i in sorted(tmp, key=naturalComponentKey):
            references, cType = tmp[i]
            references = sorted(references, key=naturalComponentKey)
            description, footprint, manufacturer, partNumber, notes, solderType = cType
            if solderType is None:
                solderType = types[references[0]]
            writer.writerow([item_no, ",".join(references),
                             len(references) * nBoards, manufacturer,
                             partNumber, description, footprint,
                             solderType, notes])
            item_no += 1

def layerName(layer):
    if layer == pcbnew.F_Cu:
        return "TopLayer"
    if layer == pcbnew.B_Cu:
        return "BottomLayer"
    raise RuntimeError(f"Got component with invalid layer {layer}")

def footprintName(footprint):
    name = str(footprint.GetFPID().GetLibItemName())
    if name[:2] == 'R_' or name[:2] == 'C_' or name[:2] == 'L_':
        name = name.split(sep='_')[1]
    return name


def collectPosData_kameda(board, correctionFields, posFilter=lambda x : True,
                   footprintX=defaultFootprintX, footprintY=defaultFootprintY, bom=None,
                   correctionFile=None):
    """
    Extract position data of the footprints.

    "Designator","Comment","Layer","Footprint","Center-X(mm)","Center-Y(mm)","Rotation","Description"
    """
    if bom is None:
        bom = {}
    else:
        bom = { getReference(comp): comp for comp in bom }

    correctionPatterns = []
    if correctionFile is not None:
        correctionPatterns = readCorrectionPatterns(correctionFile)

    footprints = []
    placeOffset = board.GetDesignSettings().GetAuxOrigin()
    for footprint in board.GetFootprints():
        if excludeFromPos(footprint):
            continue
        if posFilter(footprint) and footprint.GetReference() in bom:
            footprints.append(footprint)
    
    def getCompensation(footprint):
        if footprint.GetReference() not in bom:
            return 0, 0, 0
        field = None
        for fieldName in correctionFields:
            field = getField(bom[footprint.GetReference()], fieldName)
            if field is not None:
                break
        if field is None or field == "":
            return applyCorrectionPattern(
                correctionPatterns,
                footprint)
        try:
            return parseCompensation(field)
        except FormatError as e:
            raise FormatError(f"{footprint.GetReference()}: {e}")

    def getDescription(footprint):
        if footprint.GetReference() not in bom:
            return ""
        description = getField(bom[footprint.GetReference()], 'Description')
        if description is None:
            description = ""
        return description


    return [(footprint.GetReference(),
             footprint.GetValue(),
             layerName(footprint.GetLayer()),
             footprintName(footprint),
             footprintX(footprint, placeOffset, getCompensation(footprint)),
             footprintY(footprint, placeOffset, getCompensation(footprint)),
             footprintOrientation(footprint, getCompensation(footprint)),
             getDescription(footprint)) for footprint in footprints]

def posDataToFile_kameda(posData, filename):
    now = dt.now()

    with open(filename, "w", newline="", encoding="utf-8") as csvfile:
        csvfile.writelines(
            'Kicad/Kikit Pick and Place Locations do jeito que o Kameda gosta\n'
            f'{filename}\n\n'
            '============================================================================\n'
            'File Design Information:\n'
            f'Date: {now.strftime("%d/%m/%Y")}\n'
            f'Time: {now.strftime("%H:%M")}\n'
            'Revision: 01\n'
            'Variants: no variants\n'
            'Units used: mm\n'
            '\n\n'
        )
        writer = csv.writer(csvfile, quoting=csv.QUOTE_ALL)
        # writer.writerow(["File Design Information:"])
        # writer.writerow(["Units used: mm"])
        writer.writerow(["Designator", "Comment", "Layer", "Footprint", "Center-X(mm)", "Center-Y(mm)", "Rotation", "Description"])
        for line in sorted(posData, key=lambda x: naturalComponentKey(x[0])):
            line = list(line)
            for i in [4, 5, 6]:
                line[i] = f"{line[i]:.4f}" # Most Fab houses expect only 2 decimal digits
            writer.writerow(line)


def exportKameda(board, outputdir, assembly, schematic, ignore,
                 manufacturer, partnumber, description, notes, soldertype,
                 footprint, corrections, correctionpatterns, missingerror, nboards, nametemplate, drc):
    """
    Prepare fabrication files for Kamedatech including their assembly service
    """
    ensureValidBoard(board)
    loadedBoard = pcbnew.LoadBoard(board)

    if drc:
        ensurePassingDrc(loadedBoard)

    refsToIgnore = parseReferences(ignore)
    removeComponents(loadedBoard, refsToIgnore)
    Path(outputdir).mkdir(parents=True, exist_ok=True)

    gerberdir = os.path.join(outputdir, "gerber")
    shutil.rmtree(gerberdir, ignore_errors=True)
    gerberImpl(board, gerberdir, settings=exportSettingsPcbway)

    archiveName = expandNameTemplate(nametemplate, "gerbers", loadedBoard)
    shutil.make_archive(os.path.join(outputdir, archiveName), "zip", outputdir, "gerber")

    if not assembly:
        return
    if schematic is None:
        raise RuntimeError("When outputing assembly data, schematic is required")

    ensureValidSch(schematic)


    components = extractComponents(schematic)
    correctionFields    = [x.strip() for x in corrections.split(",")]
    manufacturerFields  = [x.strip() for x in manufacturer.split(",")]
    partNumberFields    = [x.strip() for x in partnumber.split(",")]
    descriptionFields   = [x.strip() for x in description.split(",")]
    notesFields         = [x.strip() for x in notes.split(",")]
    typeFields          = [x.strip() for x in soldertype.split(",")]
    footprintFields     = [x.strip() for x in footprint.split(",")]
    addVirtualToRefsToIgnore(refsToIgnore, loadedBoard)
    bom = collectBom(components, manufacturerFields, partNumberFields,
                     descriptionFields, notesFields, typeFields,
                     footprintFields, refsToIgnore)

    missingFields = False
    for type, references in bom.items():
        _, _, manu, partno, _, _ = type
        if not manu or not partno:
            missingFields = True
            for r in references:
                print(f"WARNING: Component {r} is missing manufacturer and/or part number")
    if missingFields and missingerror:
        sys.exit("There are components with missing ordercode, aborting")

    posData = collectPosData_kameda(loadedBoard, correctionFields, bom=components, correctionFile=correctionpatterns)
    posDataToFile_kameda(posData, os.path.join(outputdir, expandNameTemplate(nametemplate, "pos", loadedBoard) + ".csv"))
    types = collectSolderTypes(loadedBoard)
    bomToCsv(bom, os.path.join(outputdir, expandNameTemplate(nametemplate, "bom", loadedBoard) + ".csv"), nboards, types)
