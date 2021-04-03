#ifndef FD_WRITE_XLSX_HPP
#define FD_WRITE_XLSX_HPP

#include <cstring>
#include <stdexcept>
#include <string>
#include <variant>
#include <vector>
#include <zip.h>

#define FD_WRITE_XLSX_SHOW(arg) std::cout << #arg << '{' << (arg) << '}' << std::endl;

namespace fd_write_xlsx {

typedef std::string str_t;

// Representation of a cell: std::variant of string, int64_t and double. A int64_t is choosen for
// int: the size is like the size of a double.
typedef std::variant<str_t, int64_t, double> cell_t;

// Type of the sheet using by the write function.
typedef std::vector<std::vector<cell_t>> sheet_t;

// Write the contents of the “sheet” sheet into a Microsoft xlsx workbook and save it in the
// “xlsx_file_name” file.
void
write(char const* const xlsx_file_name, sheet_t const& sheet);
void
write(str_t const& xlsx_file_name, sheet_t const& sheet);
} // namespace fd_write_xlsx
#endif // FD_WRITE_XLSX_HPP
