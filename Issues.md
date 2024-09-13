# XPath version

The Python libxml XPath implementation is only XPath 1.0

Fails parsing fully both nodes.csv and fields.csv files

Example Fix:

cac:ProcurementLegislationDocumentReference[not(cbc:ID/text()=('CrossBorderLaw','LocalLegalBasis'))]

Modified to

cac:ProcurementLegislationDocumentReference[not(cbc:ID/text()='CrossBorderLaw')]
