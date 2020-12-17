/**
* Create an XSSFSheet from an existing sheet in the XSSFWorkbook.
* The cloned sheet is a deep copy of the original but with a new given
* name.
*
* @param sheetNum The index of the sheet to clone
* @param newName The name to set for the newly created sheet
* @return XSSFSheet representing the cloned sheet.
* @throws IllegalArgumentException if the sheet index or the sheet
* name is invalid
* @throws POIXMLException if there were errors when cloning
*/
public XSSFSheet cloneSheet(int sheetNum, String newName) {
validateSheetIndex(sheetNum);
XSSFSheet srcSheet = sheets.get(sheetNum);
newName = getNewName(newName);
XSSFSheet clonedSheet = createSheet(newName);
// copy sheet's relations
copySheetRelations(clonedSheet, srcSheet);
return clonedSheet;
}
private String getNewName(String newName) {
if (newName == null) {
String srcName = srcSheet.getSheetName();
newName = getUniqueSheetName(srcName);
} else {
validateSheetName(newName);
}
}
private void addXMLRelationPart(RelationPart relationPart,
XSSFDrawing xssfdDrawing, XSSFSheet clonedSheet) {
POIXMLDocumentPart documentPart = relationPart.getDocumentPart();
// do not copy the drawing relationship, it will be re-created
if(documentPart instanceof XSSFDrawing) {
xssfdDrawing = (XSSFDrawing)poiXMLDocumentPart;
continue;
}
addRelation(documentPart, clonedSheet);
}
private void legacyDrawingOrPageSetup(CTWorksheet ctWorksheet) {
if(ctWorksheet.isSetLegacyDrawing()) {
logger.log(POILogger.WARN, "Cloning sheets with comments is not yet supported.");
ctWorksheet.unsetLegacyDrawing();
}
if (ctWorksheet.isSetPageSetup()) {
logger.log(POILogger.WARN, "Cloning sheets with page setup is not yet supported.");
ctWorksheet.unsetPageSetup();
}
}
private void copySheetRelations(XSSFSheet clonedSheet, XSSFSheet srcSheet) {
List<RelationPart> relationParts = srcSheet.getRelationParts();
// if the sheet being cloned has a drawing then remember it and re-create it too
XSSFDrawing xssfdDrawing = null;
for(RelationPart relationPart : relationParts) {
addXMLRelationPart(relationPart, xssfdDrawing, clonedSheet);
}
try {
for(PackageRelationship packageRelationship :
srcSheet.getPackagePart().getRelationships()) {
if (packageRelationship.getTargetMode() == TargetMode.EXTERNAL) {
clonedSheet.getPackagePart().addExternalRelationship
(packageRelationship.getTargetURI().toASCIIString(),

packageRelationship.getRelationshipType(),
packageRelationship.getId());

}
}
} catch (InvalidFormatException e) {
throw new POIXMLException("Failed to clone sheet", e);
}
WriteTheClonedSheet(clonedSheet, srcSheet);
CTWorksheet ctWorksheet = clonedSheet.getCTWorksheet();
legacyDrawingOrPageSetup(ctWorksheet)
clonedSheet.setSelected(false);
// clone the sheet drawing along with its relationships
cloneDrawingsWithRelations(clonedSheet, xssfdDrawing, ctWorksheet);
}
private cloneDrawingsWithRelations(XSSFSheet clonedSheet, XSSFDrawing xssfdDrawing,
CTWorksheet ctWorksheet) {
if (xssfdDrawing != null) {
if(ctWorksheet.isSetDrawing()) {
// unset the existing reference to the drawing,
// so that subsequent call of clonedSheet.createDrawingPatriarch() will create a
new one
ctWorksheet.unsetDrawing();
}
XSSFDrawing clonedDrawing = clonedSheet.createDrawingPatriarch();
// copy drawing contents
clonedDrawing.getCTDrawing().set(xssfdDrawing.getCTDrawing());
clonedDrawing = clonedSheet.createDrawingPatriarch();
// Clone drawing relations
List<RelationPart> srcRelations = srcSheet.createDrawingPatriarch().getRelationParts();
for (RelationPart relationPart : srcRelations) {
addRelation(relationPart, clonedDrawing);
}
}
}
private void WriteTheClonedSheet(XSSFSheet clonedSheet, XSSFSheet srcSheet) {
try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
srcSheet.write(out);
try (ByteArrayInputStream bis = new ByteArrayInputStream(out.toByteArray())) {
clonedSheet.read(bis);
}
} catch (IOException e){
throw new POIXMLException("Failed to clone sheet", e);
}
}