import 'dart:io';
import 'package:excel/excel.dart';
import 'package:docx_template/docx_template.dart';

///
/// Read file template.docx, produce it and save
///
void main() async {
  final f = File("lib/template.docx");
  final docx = await DocxTemplate.fromBytes(await f.readAsBytes());

  int num = 2;

  while (num < 10) {
    var file = "lib/template.xlsx";
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    var sheet = excel['Sheet1'];

    String str = num.toString();
    var cell0001 = sheet.cell(CellIndex.indexByString('A$str'));
    var cell0002 = sheet.cell(CellIndex.indexByString('B$str'));
    var cell0003 = sheet.cell(CellIndex.indexByString('C$str'));
    var cell0004 = sheet.cell(CellIndex.indexByString('D$str'));
    var cell0005 = sheet.cell(CellIndex.indexByString('E$str'));
    var cell0006 = sheet.cell(CellIndex.indexByString('F$str'));
    var cell0007 = sheet.cell(CellIndex.indexByString('G$str'));
    var cell0008 = sheet.cell(CellIndex.indexByString('H$str'));
    var cell0009 = sheet.cell(CellIndex.indexByString('I$str'));
    var cell0010 = sheet.cell(CellIndex.indexByString('J$str'));
    var cell0011 = sheet.cell(CellIndex.indexByString('K$str'));
    var cell0012 = sheet.cell(CellIndex.indexByString('L$str'));
    var cell0013 = sheet.cell(CellIndex.indexByString('M$str'));
    var cell0014 = sheet.cell(CellIndex.indexByString('N$str'));
    var cell0015 = sheet.cell(CellIndex.indexByString('O$str'));
    var cell0016 = sheet.cell(CellIndex.indexByString('P$str'));
    var cell0017 = sheet.cell(CellIndex.indexByString('Q$str'));
    var cell0018 = sheet.cell(CellIndex.indexByString('R$str'));
    var cell0019 = sheet.cell(CellIndex.indexByString('S$str'));
    var cell0020 = sheet.cell(CellIndex.indexByString('T$str'));
    var cell0021 = sheet.cell(CellIndex.indexByString('U$str'));
    var cell0022 = sheet.cell(CellIndex.indexByString('V$str'));
    var cell0023 = sheet.cell(CellIndex.indexByString('W$str'));
    var cell0024 = sheet.cell(CellIndex.indexByString('X$str'));
    var cell0025 = sheet.cell(CellIndex.indexByString('Y$str'));
    var cell0026 = sheet.cell(CellIndex.indexByString('Z$str'));
    var cell0027 = sheet.cell(CellIndex.indexByString('AA$str'));

    String text0001 = cell0001.value.toString();
    String text0002 = cell0002.value.toString();
    String text0003 = cell0003.value.toString();
    if (text0003 != "") {
      text0003 = "*$text0003";
    }
    String text0004 = cell0004.value.toString();
    if (text0004 != "") {
      text0004 = "***$text0004";
    }
    String text0005 = cell0005.value.toString();
    if (text0005 != "") {
      text0005 = "**$text0005";
    }
    String text0006 = "[정답] ${cell0006.value}";
    String text0007 = "[소재] ${cell0007.value}";
    String text0008 = "[해석] ${cell0008.value}";
    String text0009 = "[해설] ${cell0009.value}";
    String text0010 = cell0010.value.toString();
    if (text0010 != "") {
      text0010 = "□$text0010";
    }
    String text0011 = cell0011.value.toString();
    String text0012 = cell0012.value.toString();
    if (text0012 != "") {
      text0012 = "□$text0012";
    }
    String text0013 = cell0013.value.toString();
    String text0014 = cell0014.value.toString();
    if (text0014 != "") {
      text0014 = "□$text0014";
    }
    String text0015 = cell0015.value.toString();
    String text0016 = cell0016.value.toString();
    if (text0016 != "") {
      text0016 = "□$text0016";
    }
    String text0017 = cell0017.value.toString();
    String text0018 = cell0018.value.toString();
    if (text0018 != "") {
      text0018 = "□$text0018";
    }
    String text0019 = cell0019.value.toString();
    String text0020 = cell0020.value.toString();
    if (text0020 != "") {
      text0020 = "□$text0020";
    }
    String text0021 = cell0021.value.toString();
    String text0022 = cell0022.value.toString();
    if (text0022 != "") {
      text0022 = "□$text0022";
    }
    String text0023 = cell0023.value.toString();
    String text0024 = cell0024.value.toString();
    if (text0024 != "") {
      text0024 = "□$text0024";
    }
    String text0025 = cell0025.value.toString();
    String text0026 = cell0026.value.toString();
    if (text0026 != "") {
      text0026 = "□$text0026";
    }
    String text0027 = cell0027.value.toString();

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
      ..add(TextContent("0027", text0027));

    final docGenerated = await docx.generate(content);
    final fileGenerated = File('result/${str}generated.docx');
    if (docGenerated != null) await fileGenerated.writeAsBytes(docGenerated);

    num++;
  }
}
