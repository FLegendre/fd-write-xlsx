# fd-write-xlsx
A fast and dirty C++ library to create a xlsx workbook.

This is an header only library : put the `fd-write-xlsx-header-only.hpp` file in a appropriate place and use it as
```C++
#include "fd-write-xlsx-header-only.hpp"

int
main()
{
  // The namespace defines “sheet_t” as the sheet type (a std::vector of std::vector).
  fd_write_xlsx::sheet_t sheet;

  // Populate the sheet with 3 rows...
  {
    using namespace fd_write_xlsx;
    // The namespace defines “cell_t” as the cell type (a std::variant).
    sheet.push_back({ cell_t{ "a" }, cell_t{ "b" }, cell_t{ "c" } });
    // To insert a integer, you need an long integer.
    sheet.push_back({ cell_t{ 1l }, cell_t{ 2l }, cell_t{ 3l } });
    sheet.push_back({ cell_t{ 1.1 }, cell_t{ 2.2 }, cell_t{ 3.3 } });
  }

  // Then, write the sheet as a Microsoft xlsx workbook.
  fd_write_xlsx::write("test.xlsx", sheet);

  return 0;
}
```

This library depends on the libzip library : https://libzip.org/.
