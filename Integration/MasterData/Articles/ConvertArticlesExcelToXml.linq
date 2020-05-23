<Query Kind="Program">
  <NuGetReference>LinqToExcel</NuGetReference>
  <Namespace>LinqToExcel</Namespace>
</Query>

// Purpose: Converts a well defined Excel with article master data into an XML file.
// Input: Articles.xls
//
// Dependendencies:
// https://www.linqpad.net/LINQPad5.aspx: choco install linqpad5.install
// LinqToExcel: Excel access with LinkToExcel: ExcelQueryFactory, source of example: https://www.c-sharpcorner.com/article/linq-to-excel-in-action/

public static string FMGDATAEXCHANGE_VERSION = "1.0";
public static string ARTICLES_VERSION = "1.6.7";
void Main() {
	var dir = Path.GetDirectoryName(Util.CurrentQueryPath);
	String filePathLoadExcel = Path.Combine(dir, @"Articles.xls");
	Console.WriteLine(filePathLoadExcel);
	String filePathXmlExport = Path.Combine(dir, "Articles_Export.xml");

	ConnexionExcel ConxObject = new ConnexionExcel(filePathLoadExcel);
	//Query a worksheet with a header row  
	var artikelstamm = from a in ConxObject.UrlConnexion.Worksheet<ImportArtikel>("ArticleMasterData") select a;
	//artikelstamm.Dump();
	List<XElement> items = new List<XElement>();
	foreach (var row in artikelstamm) {
		items.Add(new XElement("Article",
			new XElement("Name", row.Artikelnummer1),
			new XElement("Name2", ""),
			new XElement("Description", row.Bezeichnung1),
			new XElement("Description2", row.Bezeichnung2),
			new XElement("Description3", row.Bezeichnung3),
			new XElement("Code", ""),
			new XElement("Dimension",
				new XElement("Length", row.Laenge),
				new XElement("Width", row.Breite),
				new XElement("Height", row.Dicke),
				new XElement("Unit", "mm")
			),
			new XElement("MaterialGroup",
				new XElement("Name", row.Materialgruppe)
			),
			new XElement("Material", new XAttribute("Id", "0"),
				new XElement("Name", row.Materialgruppe),
				new XElement("SpecificWeight", row.Spezgewicht),
				new XElement("UnitOfSpecificWeight", "kg/dm3")
			),
			new XElement("TotalQuantity", "0"),
			new XElement("StockUnit", "0"),
			new XElement("TotalWeight", "0"),
			new XElement("MinimumUnclaimedQuantity", "0")
			)
		);
	}

	var xmlDoc = new XDocument(
			 new XDeclaration("1.0", "utf-8", null),
			 new XElement("FmgDataExchange", new XAttribute("Version", FMGDATAEXCHANGE_VERSION),
			 new XElement("Header",
			 	new XElement("CreatedDateTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")),
				new XElement("LanguageCode", "de"),
				new XElement("Source", new XAttribute("Id", "PPS"),
					new XElement("Description", "")
				),
				new XElement("Destination", new XAttribute("Id", "LVS"),
					new XElement("Description", "FMG LVS")
				)
				),
			 new XElement("Articles", new XAttribute("Version", ARTICLES_VERSION),
			 	items)));
	xmlDoc.Save(filePathXmlExport);
	xmlDoc.Dump();	
}

public class ImportArtikel {
	public string Artikelnummer1 { get; set; }
	public string Bezeichnung1 { get; set; }
	public string Bezeichnung2 { get; set; }
	public string Bezeichnung3 { get; set; }
	public double Laenge { get; set; }
	public double Breite { get; set; }
	public double Dicke { get; set; }
	public double Spezgewicht { get; set; }
	public string Materialgruppe { get; set; }
}

public class ConnexionExcel {
	public string _pathExcelFile;
	public ExcelQueryFactory _urlConnexion;
	public ConnexionExcel(string path) 	{
		this._pathExcelFile = path;
		this._urlConnexion = new ExcelQueryFactory(_pathExcelFile);
	}
	public string PathExcelFile => _pathExcelFile;
	public ExcelQueryFactory UrlConnexion => _urlConnexion;
}