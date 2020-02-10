#include <iostream>
#include <iomanip>
#include <dmxlsx.h>
#include "dmutil.h"

using namespace std;
/*
 * TODO: Sheet iterator
 * TODO: Handling of named ranges
 * TODO: Column/Row iterators
 * TODO: correct copy/move operations for all classes
 * TODO: Find a way to handle overwriting of shared formulas.
 * TODO: Handling of Cell formatting
 * TODO: Handle chartsheets
 * TODO: Update formulas when changing Sheet Name.
 * TODO: Get vector for a Row or Column.
 * TODO: Conditional formatting
 */

int main() {
    try
    {
        XLDocument doc;
        doc.OpenDocument(DMGetRootPath() + PATH_DELIMITER + "1monster.xlsx");
        auto wks = doc.Workbook().Worksheet("def");

        for (int i = 1; i <= wks.RowCount(); ++i)
        {
            auto row = wks.Row(i);
            for (int j = 1; j <= row.CellCount(); ++j)
            {
                auto cell = row.Cell(j);
                std::cout << cell.Value().AsString() << "|";
            }
            std::cout << std::endl;
        }
    }
    catch (std::runtime_error& e)
    {
        std::cout << e.what() << std::endl;
    }


    return 0;
}
