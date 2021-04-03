#include "fd-write-xlsx-header-only.hpp"

int
main()
{
	// “sheet_t” is a std::vector of std::vectors of “cell_t” instances and “cell_t” is a
	// std::variant<std::string, int64_t, double>.

	fd_write_xlsx::sheet_t sheet;

	// Populate the sheet with 3 rows...
	{
		sheet.push_back({ { "a" }, { "b" }, { "c" } });
		// To insert a integer, you need an long integer.
		sheet.push_back({ { 1l }, { 2l }, { 3l } });
		sheet.push_back({ { 1.1 }, { 2.2 }, { 3.3 } });
	}

	// Then, write the sheet as a Microsoft xlsx workbook.
	fd_write_xlsx::write("test.xlsx", sheet);

	return 0;
}
