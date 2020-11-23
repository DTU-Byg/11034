import ifcopenshell
import xlsxwriter # https://xlsxwriter.readthedocs.io/tutorial01.html#tutorial1

workbook = xlsxwriter.Workbook('output/test2.xlsx')
model = ifcopenshell.open("model/Duplex_A_20110907.ifc")

def makeASheet(ifcType):
    sheet = workbook.add_worksheet(ifcType)
    row =1
    for entity in model.by_type(ifcType):
        sheet.write(row,0,str(entity.Name))
        row+=1

makeASheet('IfcSlab')
makeASheet('IfcWall')
makeASheet('IfcCovering')
makeASheet('IfcBeam')

workbook.close()