/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author loay
 */
private void copySheetRelations( XSSFSheet srcSheet) {   // i divide this function for copy sheet's relations

        List<RelationPart> rels = srcSheet.getRelationParts();
        // if the sheet being cloned has a drawing then remember it and re-create it too
        XSSFDrawing xssfdrwa = null;
        for(RelationPart relationpart : rels) {
            POIXMLDocumentPart r = relationpart.getDocumentPart();
            // do not copy the drawing relationship, it will be re-created
            if(r instanceof XSSFDrawing) {
                xssfdrwa = (XSSFDrawing)r;
                continue;
            }}}
private String ValidateName(String newName){  // i divide this function for validate name and set new name 
 if (newName == null) {
            String srcName = srcSheet.getSheetName();
            newName = getUniqueSheetName(srcName);
            return newName;
        } else {
         validateSheetName(newName);
           return newName;

       }}

public XSSFSheet cloneSheet(int sheetNumber, String newName) {
        validateSheetIndex(sheetNumber);
        XSSFSheet srcSheet = sheets.get(sheetNumber);
         newName= ValidateName(newName);


        XSSFSheet clonedSheet = createSheet(newName);  //new sheet to clone original 

       
copySheetRelations(srcSheet);

            addRelation(relationpart, clonedSheet);

        

        try {
            for(PackageRelationship packagerelation : srcSheet.getPackagePart().getRelationships()) {
                if (packagerelation.getTargetMode() == TargetMode.EXTERNAL) {
                    clonedSheet.getPackagePart().addExternalRelationship
                            (packagerelation.getTargetURI().toASCIIString(), packagerelation.getRelationshipType(), packagerelation.getId());
                }
            }
        } catch (InvalidFormatException e) {
            throw new POIXMLException("Failed to clone sheet", e);
        }


        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            srcSheet.write(out);
            try (ByteArrayInputStream bis = new ByteArrayInputStream(out.toByteArray())) {
                clonedSheet.read(bis);
            }
        } catch (IOException e){
            throw new POIXMLException("Failed to clone sheet", e);
        }
        CTWorksheet ct = clonedSheet.getCTWorksheet();
        if(ct.isSetLegacyDrawing()) {
            logger.log(POILogger.WARN, "Cloning sheets with comments is not yet supported.");
            ct.unsetLegacyDrawing();
        }
        if (ct.isSetPageSetup()) {
            logger.log(POILogger.WARN, "Cloning sheets with page setup is not yet supported.");
            ct.unsetPageSetup();
        }

        clonedSheet.setSelected(false);

        // clone the sheet drawing along with its relationships
        if (xssfdrwa != null) {
            if(ct.isSetDrawing()) {
                // unset the existing reference to the drawing,
                // so that subsequent call of clonedSheet.createDrawingPatriarch() will create a new one
                ct.unsetDrawing();
            }
            XSSFDrawing clonedDg = clonedSheet.createDrawingPatriarch();
            // copy drawing contents
            clonedDg.getCTDrawing().set(xssfdrwa.getCTDrawing());

            clonedDg = clonedSheet.createDrawingPatriarch();

            // Clone drawing relations
            List<RelationPart> srcRels = srcSheet.createDrawingPatriarch().getRelationParts();
            for (RelationPart relationpart : srcRels) {
                addRelation(relationpart, clonedDg);
            }
        }
        return clonedSheet;
    }
}
