import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'dart:io';

void main() => runApp(const TodoApp());

class TodoApp extends StatelessWidget {
  const TodoApp({Key? key}) : super(key: key);

  @override
  Widget build(BuildContext context) {
    return const SheetList();
  }
}

class SheetList extends StatefulWidget {
  const SheetList({super.key});

  @override
  createState() => _SheetListState();
}

class XLSX {
  late String title;
  late int maxCols;
  late int maxRows;
  late List<String> headers;
  late List<List<String>> _data;
  late List<List<String>> displayData;

  XLSX(Sheet sheet) {
    title = sheet.sheetName;
    maxCols = sheet.maxCols;
    maxRows = sheet.maxRows;
    headers = sheet.rows[0]
        .map((Data? e) => e != null ? e.value.toString() : "Null")
        .toList();
    _data = sheet.rows.getRange(1, sheet.rows.length).map((List<Data?> row) {
      return row
          .map((Data? e) => e != null ? e.value.toString() : "Null")
          .toList();
    }).toList();
    displayData = _data;
  }

  void sort(int colIndex, bool descending) {
    bool isNumeric = double.tryParse(displayData[0][colIndex]) != null;
    if (isNumeric) {
      displayData.sort((a, b) =>
          double.parse(a[colIndex]).compareTo(double.parse(b[colIndex])));
    } else {
      displayData.sort((a, b) => a[colIndex].compareTo(b[colIndex]));
    }
    if (descending) displayData = displayData.reversed.toList();
  }

  void filter(int colIndex, String str) {
    if (str == "") displayData = _data;
    displayData = _data
        .where((List<String> row) =>
            row[colIndex].toLowerCase().contains(str.toLowerCase()))
        .toList();
  }
}

class _SheetListState extends State<SheetList> {
  List<XLSX>? tables;
  int _sortColumnIndex = 0;
  bool _sortAscending = false;
  final _textController = TextEditingController();
  String str = '';
  int _selected = 0;

  Widget _buildExcel() {
    return TabBarView(
        children: Iterable<int>.generate(tables!.length)
            .map((index) => _buildSheet(index))
            .toList());
  }

  Widget _buildFilterBar(int index) {
    return Row(
      children: [
        Expanded(
          child: Padding(
            padding: const EdgeInsets.all(10.0),
            child: TextField(
              controller: _textController,
              decoration: InputDecoration(
                  hintText: 'Search here...',
                  suffixIcon: IconButton(
                    onPressed: () {
                      _textController.clear();
                      setState(() {
                        tables![index].filter(_selected, "");
                      });
                    },
                    icon: const Icon(Icons.clear),
                  )),
              onChanged: (String newText) {
                setState(() {
                  tables![index].filter(_selected, newText);
                });
              },
            ),
          ),
        ),
        Padding(
          padding: const EdgeInsets.all(10.0),
          child: DropdownButton(
            items: Iterable<int>.generate(tables![index].headers.length)
                .map((int i) {
              return DropdownMenuItem<int>(
                value: i,
                child: Text(tables![index].headers[i]),
              );
            }).toList(),
            value: _selected,
            onChanged: dropdownCallback,
          ),
        )
      ],
    );
  }

  Widget _buildSheet(int index) {
    return Column(children: [
      _buildFilterBar(index),
      Expanded(
        child: SingleChildScrollView(
            scrollDirection: Axis.vertical,
            child: SingleChildScrollView(
                scrollDirection: Axis.horizontal,
                child: DataTable(
                    sortColumnIndex: _sortColumnIndex,
                    sortAscending: _sortAscending,
                    columns: tables![index]
                        .headers
                        .map((String row) => DataColumn(
                            onSort: (columnIndex, sortAscending) {
                              setState(() {
                                if (_sortColumnIndex == columnIndex) {
                                  _sortAscending = !_sortAscending;
                                } else {
                                  _sortColumnIndex = columnIndex;
                                  _sortAscending = false;
                                }
                                tables![index]
                                    .sort(columnIndex, _sortAscending);
                              });
                            },
                            label: Text(row,
                                style: const TextStyle(
                                    fontSize: 20,
                                    fontWeight: FontWeight.bold,
                                    color: Colors.black))))
                        .toList(),
                    rows: tables![index]
                        .displayData
                        .map((List<String> row) => DataRow(
                            cells: row
                                .map((String e) => DataCell(Text(e)))
                                .toList()))
                        .toList()))),
      )
    ]);
  }

  PreferredSizeWidget _buildTabs(Iterable<String> titles) {
    return TabBar(tabs: titles.map((title) => Tab(text: title)).toList());
  }

  void dropdownCallback(int? selectedValue) {
    if (selectedValue is int) {
      setState(() {
        _selected = selectedValue;
      });
    }
  }

  _pick() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles();
    if (result != null) {
      setState(() {
        var bytes = File(result.files.single.path!).readAsBytesSync();
        Excel excel = Excel.decodeBytes(bytes);
        tables = excel.tables.entries
            .where((entry) => entry.value.maxCols > 0)
            .map((entry) => XLSX(entry.value))
            .toList();
      });
    }
  }

  @override
  Widget build(BuildContext context) {
    return (tables != null
        ? MaterialApp(
            home: DefaultTabController(
                length: tables!.length,
                child: Scaffold(
                    appBar: PreferredSize(
                      preferredSize: const Size.fromHeight(kToolbarHeight),
                      child: Container(
                        color: Colors.blue,
                        child: SafeArea(
                          child: _buildTabs(
                              tables!.map((XLSX sheet) => sheet.title)),
                        ),
                      ),
                    ),
                    body: _buildExcel(),
                    floatingActionButton: FloatingActionButton(
                        onPressed: _pick,
                        tooltip: 'Add task',
                        child: const Icon(Icons.add)))))
        : MaterialApp(
            home: Scaffold(
                floatingActionButton: FloatingActionButton(
                    onPressed: _pick,
                    tooltip: 'Add task',
                    child: const Icon(Icons.add)))));
  }
}
