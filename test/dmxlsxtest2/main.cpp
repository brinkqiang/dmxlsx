#include <iostream>
#include <chrono>
#include <iomanip>
#include <dmxlsx.h>

using namespace std;

int main() {

    XLDocument doc;
    doc.CreateDocument("./MyTest.xlsx");
    auto wbk = doc.Workbook();

    wbk.AddWorksheet("MySheet01");    // Append new sheet
    wbk.AddWorksheet("MySheet02", 1); // Prepend new sheet
    wbk.AddWorksheet("MySheet03", 1); // Prepend new sheet
    wbk.AddWorksheet("MySheet04", 2); // Insert new sheet
    wbk.MoveSheet("Sheet1", 2);       // Move Sheet1 to second place
    wbk.DeleteSheet("MySheet01");

    wbk.Worksheet("Sheet1").Row(1);

    for (const auto& name : wbk.WorksheetNames())
    {
        cout << name << ": " << wbk.IndexOfSheet(name) << endl;
    }

    cout << endl;

    for (auto iter = 1; iter <= wbk.SheetCount(); ++iter)
    {
        cout << iter << ": " << wbk.Sheet(iter).Name() << endl;        
    }

    doc.SaveDocument();

    XLDocument doc2;
    doc2.OpenDocument("./MyTest.xlsx");
    auto wbk2 = doc2.Workbook();

    for (const auto& name : wbk2.WorksheetNames())
    {
        cout << name << ": " << wbk2.IndexOfSheet(name) << endl;
    }

    cout << endl;

    for (auto iter = 1; iter <= wbk2.SheetCount(); ++iter)
    {
        cout << iter << ": " << wbk2.Sheet(iter).Name() << endl;
    }

    return 0;
}
