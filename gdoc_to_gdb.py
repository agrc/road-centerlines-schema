import gspread
import arcpy
from oauth2client.service_account import ServiceAccountCredentials
from time import strftime

uniqueRunNum = strftime("%Y%m%d_%H%M%S")
# Values from the spreadsheet mapped to the geoprocessing tool parameter string
domainTranslations = {
                      'CodedValue': 'CODED',
                      'String': 'TEXT',
                      'DefaultValue': 'DEFAULT',
                      'null': None,
                      'SmallInteger': 'SHORT',
                      'TBD': 'TEXT'
                      }


class Domain(object):
    """Store data used for GDB domain creation."""

    def __init__(self, domainName, domainType, fieldType, mergePolicy, splitPolicy, description, owner):
        """Constructor."""
        self.domainName = domainName
        self.domainType = domainType
        self.fieldType = fieldType
        self.mergePolicy = mergePolicy
        self.splitPolicy = splitPolicy
        self.description = description
        self.owner = owner
        self.codedValues = {}

    def addCodedValue(self, code, value):
        """Store coded value for the domain."""
        self.codedValues[code] = value

    def addToWorkspace(self, workspace):
        """Add this domain to the GDB workspace."""
        arcpy.CreateDomain_management (workspace,
                                       self.domainName,
                                       self.description,
                                       self.fieldType,
                                       self.domainType,
                                       self.splitPolicy,
                                       self.mergePolicy)
        for code in self.codedValues:
            arcpy.AddCodedValueToDomain_management (workspace,
                                                    self.domainName,
                                                    code,
                                                    self.codedValues[code])


class Field(object):
    """Store data used for field creation."""

    def __init__(self, fieldName, fieldType, fieldLength, aliasName, domainName):
        """Constructor."""
        self.fieldName = fieldName
        self.fieldType = fieldType
        self.fieldLength = fieldLength
        self.aliasName = aliasName
        self.domainName = domainName

    def addToFeatureClass(self, featureClass):
        """Add this field to the featureClass."""
        length = None
        if self.fieldLength.isdigit():
            length = int(self.fieldLength)
        arcpy.AddField_management(in_table=featureClass,
                                  field_name=self.fieldName,
                                  field_type=self.fieldType,
                                  field_length=length,
                                  field_alias=self.aliasName,
                                  field_domain=self.domainName)


def checkStrParam(spreadSheetString):
    """Check strings required for geoprocessing tool parameters."""
    return domainTranslations.get(spreadSheetString, spreadSheetString)


def getFields(fieldWorksheet, nameI, typeI, lengthI, aliasI, domainI, fieldTableRow):
    """Create a list of Field objects from the centerline worksheet."""
    fullSheetList = fieldWorkSheet.get_all_values()
    fieldRows = fullSheetList[fieldTableRow:]
    return [Field(f[nameI], checkStrParam(f[typeI]), f[lengthI], f[aliasI], f[domainI]) for f in fieldRows]


def getDomains(worksheets,
               domainNameI,
               domainTypeI,
               fieldTypeI,
               mergePolicyI,
               splitPolicyI,
               descriptionI,
               ownerI,
               codedValueHeaderI):
    """Create a list of Domain objects from all worksheets that have Domain in the title."""
    domains = []
    for ws in worksheets:
        if 'Domain' in ws.title:
            wsList = ws.get_all_values()
            d = Domain(checkStrParam(wsList[domainNameI][1]),
                       checkStrParam(wsList[domainTypeI][1]),
                       checkStrParam(wsList[fieldTypeI][1]),
                       checkStrParam(wsList[mergePolicyI][1]),
                       checkStrParam(wsList[splitPolicyI][1]),
                       checkStrParam(wsList[descriptionI][1]),
                       checkStrParam(wsList[ownerI][1]))

            codedValList = wsList[codedValueHeaderI:]
            for cv in codedValList:
                d.addCodedValue(cv[0], cv[1])

            domains.append(d)
    return domains


if __name__ == '__main__':
    scope = ['https://spreadsheets.google.com/feeds']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(r'CenterlineSchema-5db1aa340548.json', scope)
    gc = gspread.authorize(credentials)
    # spreadsheet must be shared with the email in credentials
    spreadSheet = gc.open_by_url(r"https://docs.google.com/spreadsheets/d/1QbhvmE-HEPcYM7qWGbSxh1F8FApq9LN22MkQM3fU8nE/edit#gid=811360546")
    worksheets = spreadSheet.worksheets()
    fieldWorkSheet = spreadSheet.worksheet('FC_RoadCenterlines')

    # Set up field indicies for worksheet list access
    nameI = fieldWorkSheet.find('FieldName').col - 1
    typeI = fieldWorkSheet.find('Type').col - 1
    lengthI = fieldWorkSheet.find('Length').col - 1
    aliasI = fieldWorkSheet.findall('AliasName')[1].col - 1  # AliasName is used multiple times in sheet
    domainI = fieldWorkSheet.findall('DomainName')[1].col - 1  # DomainName is used multiple times in sheet
    fieldTableRow = fieldWorkSheet.find('FieldName').row
    # Get the fields from the FC_RoadCenterlines sheet
    fields = getFields(fieldWorkSheet,
                       nameI,
                       typeI,
                       lengthI,
                       aliasI,
                       domainI,
                       fieldTableRow)
    # Set up indicies for domain worksheets
    domainNameI = 0
    domainTypeI = 1
    fieldTypeI = 2
    mergePolicyI = 3
    splitPolicyI = 4
    descriptionI = 5
    ownerI = 6
    codedValueHeaderI = 10
    # Get the domains from all Domain worksheets
    domains = getDomains(worksheets,
                         domainNameI,
                         domainTypeI,
                         fieldTypeI,
                         mergePolicyI,
                         splitPolicyI,
                         descriptionI,
                         ownerI,
                         codedValueHeaderI)

    # Create GDB for domains and output feature class
    outputGdb = arcpy.CreateFileGDB_management(r'C:\GisWork\GdocRoadsSchema',
                                               'CenterLineSchema' + uniqueRunNum)
    print 'Output GDB created'
    for d in domains:
        print 'Adding domain: {}'.format(d.domainName)
        d.addToWorkspace(outputGdb)

    # Create road centerline feature class
    outputRoadFc = arcpy.CreateFeatureclass_management(out_path=outputGdb,
                                                       out_name='RoadCenterlines',
                                                       geometry_type='POLYLINE')
    print 'Road feature class created'
    for f in fields:
        print 'Add field: {}'.format(f.fieldName)
        f.addToFeatureClass(outputRoadFc)
