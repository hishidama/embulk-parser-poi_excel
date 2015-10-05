Embulk::JavaPlugin.register_parser(
  "poi_excel", "org.embulk.parser.poi_excel.PoiExcelParserPlugin",
  File.expand_path('../../../../classpath', __FILE__))
