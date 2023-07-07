import 'dart:io';
import 'package:path/path.dart';
import 'package:excel/excel.dart';
import 'package:docx_template/docx_template.dart';

///
/// Read file template.docx, produce it and save
///
void main() async {
  final f = File("lib/template2.docx");
  final docx = await DocxTemplate.fromBytes(await f.readAsBytes());

  int num = 2;

  while (num < 10) {
    var file = "lib/template.xlsx";
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    var sheet = excel['Sheet1'];

    String str = num.toString();
    print(str);
    var cell0001 = sheet.cell(CellIndex.indexByString('A$str'));
    print(cell0001.value);
    print(cell0001.value.runtimeType);

    String text0001 = cell0001.value.toString();
    String text0002 = "Thomas Jefferson's knowledge";
    String text0003 = "reconnaissance";
    if (text0003 != "") {
      text0003 = "*$text0003";
    }
    String text0004 = "정찰";
    String text0005 = "strain";
    if (text0005 != "") {
      text0005 = "**$text0005";
    }
    String text0006 = "품종";
    String text0007 = "undaunted";
    if (text0007 != "") {
      text0007 = "***$text0007";
    }
    String text0008 = "굴하지 않는2";
    String text0009 = "[정답]";
    String text0010 = "5";
    String text0011 = "[소재]";
    String text0012 = "동화 작가이자 삽화가인 Leo Lionni";
    String text0013 = "[해석]";
    String text0014 = "국제적으로 알려진 디자이너이자 삽화가이자 그래픽 아티스트였던 ";
    String text0015 = "[해설]";
    String text0016 = "1982년, Lionni는 파킨슨병을 진단받았지만, ";
    String text0017 = "[어휘]";
    String text0018 = "illustrator";
    if (text0018 != "") {
      text0018 = "□$text0018";
    }
    String text0019 = "삽화가";
    String text0020 = "illustrator";
    if (text0020 != "") {
      text0020 = "□$text0020";
    }
    String text0021 = "삽화가";
    String text0022 = "illustrator";
    if (text0022 != "") {
      text0022 = "□$text0022";
    }
    String text0023 = "삽화가";
    String text0024 = "illustrator";
    if (text0024 != "") {
      text0024 = "□$text0024";
    }
    String text0025 = "삽화가";
    String text0026 = "illustrator";
    if (text0026 != "") {
      text0026 = "□$text0026";
    }
    String text0027 = "삽화가";
    String text0028 = "illustrator";
    if (text0028 != "") {
      text0028 = "□$text0028";
    }
    String text0029 = "삽화가";
    String text0030 = "illustrator";
    if (text0030 != "") {
      text0030 = "□$text0030";
    }
    String text0031 = "삽화가";
    String text0032 = "illustrator";
    if (text0032 != "") {
      text0032 = "□$text0032";
    }
    String text0033 = "삽화가";
    String text0034 = "illustrator";
    if (text0034 != "") {
      text0034 = "□$text0034";
    }
    String text0035 = "삽화가";
    String text0036 = "illustrator";
    if (text0036 != "") {
      text0036 = "□$text0036";
    }
    String text0037 = "삽화가";
    String text0038 = "illustrator";
    if (text0038 != "") {
      text0038 = "□$text0038";
    }
    String text0039 = "삽화가";

    Content content = Content();
    content
      ..add(TextContent("0001", text0001))
      ..add(TextContent("0002", text0002))
      ..add(TextContent("0003", text0003))
      ..add(TextContent("0004", text0004))
      ..add(TextContent("0005", text0005))
      ..add(TextContent("0006", text0006))
      ..add(TextContent("0007", text0007))
      ..add(TextContent("0008", text0008))
      ..add(TextContent("0009", text0009))
      ..add(TextContent("0010", text0010))
      ..add(TextContent("0011", text0011))
      ..add(TextContent("0012", text0012))
      ..add(TextContent("0013", text0013))
      ..add(TextContent("0014", text0014))
      ..add(TextContent("0015", text0015))
      ..add(TextContent("0016", text0016))
      ..add(TextContent("0017", text0017))
      ..add(TextContent("0018", text0018))
      ..add(TextContent("0019", text0019))
      ..add(TextContent("0020", text0020))
      ..add(TextContent("0021", text0021))
      ..add(TextContent("0022", text0022))
      ..add(TextContent("0023", text0023))
      ..add(TextContent("0024", text0024))
      ..add(TextContent("0025", text0025))
      ..add(TextContent("0026", text0026))
      ..add(TextContent("0027", text0027))
      ..add(TextContent("0028", text0028))
      ..add(TextContent("0029", text0029))
      ..add(TextContent("0030", text0030))
      ..add(TextContent("0031", text0031))
      ..add(TextContent("0032", text0032))
      ..add(TextContent("0033", text0033))
      ..add(TextContent("0034", text0034))
      ..add(TextContent("0035", text0035))
      ..add(TextContent("0036", text0036))
      ..add(TextContent("0037", text0037))
      ..add(TextContent("0038", text0038))
      ..add(TextContent("0039", text0039));

    final docGenerated = await docx.generate(content);
    final fileGenerated = File('result/${str}generated.docx');
    if (docGenerated != null) await fileGenerated.writeAsBytes(docGenerated);

    num++;
  }
}
