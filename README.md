# fd-write-xlsx
A fast and dirty C++ library to create a xlsx workbook.

This is an header only library : put the `fd-write-xlsx.hpp` file in a appropriate place and use it as
```C++
#include "fd-write-xlsx.hpp"

int
main()
{
  // The namespace defines “table_t” as the table type (a std::vector of std::vector).
  fd_write_xlsx::table_t table;

  // Populate the table with 3 rows...
  {
    using namespace fd_write_xlsx;
    // The namespace defines “cell_t” as the cell type (a std::variant).
    table.push_back({ cell_t{ "a" }, cell_t{ "b" }, cell_t{ "c" } });
    // To insert a integer, you need an long integer.
    table.push_back({ cell_t{ 1l }, cell_t{ 2l }, cell_t{ 3l } });
    table.push_back({ cell_t{ 1.1 }, cell_t{ 2.2 }, cell_t{ 3.3 } });
  }

  // Then, write the table as a Microsoft xlsx workbook.
  fd_write_xlsx::write("test.xlsx", table);

  return 0;
}
```

This library depends on the libzip library : https://libzip.org/.
