#ifndef FD_WRITE_XLSX_HPP
#define FD_WRITE_XLSX_HPP

#include <cstring>
#include <stdexcept>
#include <string>
#include <variant>
#include <vector>
#include <zip.h>

#define FD_WRITE_XLSX_SHOW(arg) std::cout << #arg << '{' << arg << '}' << std::endl;

namespace fd_write_xlsx {

typedef std::string str_t;

class Exception : public std::exception
{
public:
	Exception(str_t const& msg)
		: msg_("fd-write-xslx library: " + msg + '.')
	{}
	const char* what() const throw() { return msg_.c_str(); }

private:
	str_t const msg_;
};

// Representation of a cell: std::variant of string, int64_t and double. A int64_t is choosen for
// int: the size is like the size of a double.
typedef std::variant<str_t, int64_t, double> cell_t;

// Type of the table using by the write function.
typedef std::vector<std::vector<cell_t>> table_t;

// Class for RAII.
struct Zip
{
	Zip(char const* const file_name)
	{
		int zip_error;
		archive_ptr_ = zip_open(file_name, ZIP_CREATE | ZIP_TRUNCATE, &zip_error);
		if (!archive_ptr_)
			throw Exception{ "unable to open the “" + str_t(file_name) + "” workbook" };
	}
	~Zip()
	{
		if (archive_ptr_)
			// Do not check the return code as is bad to throw an exception in a destructor...
			zip_close(archive_ptr_);
	}
	zip_t* archive_ptr_;
};

// Write the contents of the “table” table into a Microsoft xlsx workbook and save it in the
// “xlsx_file_name” file.
void
write(char const* const xlsx_file_name, table_t const& table)
{
	auto const zip{ Zip{ xlsx_file_name } };

	// The archive tree is
	//          _rels
	//          xl
	//          [Content_Types].xml

	// Write the “[Content_Types].xml” file.
	{
		auto const str{
			R"xxxx(<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>
)xxxx"
		};
		auto source_ptr{ zip_source_buffer(zip.archive_ptr_, str, std::strlen(str), 0) };
		if (source_ptr == nullptr) {
			zip_source_free(source_ptr);
			throw Exception{ "unable to get a buffer to add “[Content_Types].xml” file" };
		}
		if (zip_file_add(zip.archive_ptr_, "[Content_Types].xml", source_ptr, 0) < 0) {
			zip_source_free(source_ptr);
			throw Exception{ "unable to add “[Content_Types].xml” file" };
		}
	}

	// Create the “_rels” directory.
	if (zip_dir_add(zip.archive_ptr_, "_rels", 0) < 0)
		throw Exception{ "unable to add “_rels” directory" };
	// Write the “_rels/.rels” file.
	{
		auto const str{ R"xxxx(<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
)xxxx" };
		auto source_ptr{ zip_source_buffer(zip.archive_ptr_, str, std::strlen(str), 0) };
		if (source_ptr == nullptr) {
			zip_source_free(source_ptr);
			throw Exception{ "unable to get a buffer to add “_rels/.rels” file" };
		}
		if (zip_file_add(zip.archive_ptr_, "_rels/.rels", source_ptr, 0) < 0) {
			zip_source_free(source_ptr);
			throw Exception{ "unable to add “_rels/.rels” file" };
		}
	}

	// Create the “xl” directory.
	if (zip_dir_add(zip.archive_ptr_, "xl", 0) < 0)
		throw Exception{ "unable to add “xl” directory" };
	// Write the “xl/workbook.xml” file.
	{
		auto const str{
			R"xxxx(<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="sheet1" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>
)xxxx"
		};
		auto source_ptr{ zip_source_buffer(zip.archive_ptr_, str, std::strlen(str), 0) };
		if (source_ptr == nullptr) {
			zip_source_free(source_ptr);
			throw Exception{ "unable to get a buffer to add “xl/workbook.xml” file" };
		}
		if (zip_file_add(zip.archive_ptr_, "xl/workbook.xml", source_ptr, 0) < 0) {
			zip_source_free(source_ptr);
			throw Exception{ "unable to add “xl/workbook.xml” file" };
		}
	}

	// Create the “xl/_rels” directory.
	if (zip_dir_add(zip.archive_ptr_, "xl/_rels", 0) < 0)
		throw Exception{ "unable to add “xl/_rels” directory" };
	// Write the “xl/_rels/workbook.xml.rels” file.
	{
		auto const str{ R"xxxx(<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
)xxxx" };
		auto source_ptr{ zip_source_buffer(zip.archive_ptr_, str, std::strlen(str), 0) };
		if (source_ptr == nullptr) {
			zip_source_free(source_ptr);
			throw Exception{ "unable to get a buffer to add “xl/_rels/workbook.xml.rels” file" };
		}
		if (zip_file_add(zip.archive_ptr_, "xl/_rels/workbook.xml.rels", source_ptr, 0) < 0) {
			zip_source_free(source_ptr);
			throw Exception{ "unable to add “xl/_rels/workbook.xml.rels” file" };
		}
	}

	// Create the “xl/worksheets” directory.
	if (zip_dir_add(zip.archive_ptr_, "xl/worksheets", 0) < 0)
		throw Exception{ "unable to add “xl/worksheets” directory" };
	// Write the “xl/worksheets/sheet1.xml” file.
	// <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
	//   <sheetData>
	//     <row r="1">
	//       <c r="A1" t="inlineStr">
	//         <is>
	//           <t>A</t>
	//         </is>
	//       </c>
	//     </row>
	//     <row r="2">
	//       <c r="A2">
	//         <v>1</v>
	//       </c>
	//     </row>
	//     <row r="3">
	//       <c r="A3">
	//         <v>1.1</v>
	//       </c>
	//     </row>
	//   </sheetData>
	// </worksheet>
	{
		size_t constexpr buffer_size{ 1024 * 16 };
		char buffer[buffer_size];
		zip_error_t error;
		auto source_ptr{ zip_source_buffer_create(buffer, buffer_size, 0, &error) };
		if (!source_ptr)
			throw Exception{ "unable to create a buffer" };
		if (zip_source_begin_write(source_ptr) < 0)
			throw Exception{ "unable to begin write" };
		auto const write{ [&](char const* const str) {
			if (zip_source_write(source_ptr, str, std::strlen(str)) < 0)
				throw Exception{ "unable to write" };
		} };
		auto const xmlify{ [](str_t const& str) {
			str_t rvo{};
			for (auto const& c : str)
				if (c == '<')
					rvo += "&lt;";
				else
					rvo += c;
			return rvo;
		} };
		write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
		write("<sheetData>");
		for (auto const& row : table) {
			auto const str_i{ std::to_string(&row - &table[0] + 1) };
			write(("<row r=\"" + str_i + "\">").c_str());
			for (auto const& cell : row) {
				// Generate the column number as a sequence of letters : A, B, ..., AA, AB, ...
				auto const j{ &cell - &row[0] };
				auto const str_j{
					(j < 26) ? str_t(1, char('A' + j))
									 : (j < 26 + 26 * 26)
											 ? (str_t(1, char('A' + (j / 26 - 1) % 26)) + str_t(1, char('A' + j % 26)))
											 : (str_t(1, char('A' + (j / (26 * 26) - 1) % 26)) +
													str_t(1, char('A' + (j / 26 - 1) % 26)) + str_t(1, char('A' + j % 26)))
				};
				if (std::holds_alternative<str_t>(cell))
					write(("<c r=\"" + str_j + str_i + "\"  t=\"inlineStr\"><is><t>" +
								 xmlify(std::get<str_t>(cell)) + "</t></is></c>")
									.c_str());
				else if (std::holds_alternative<int64_t>(cell))
					write(("<c r=\"" + str_j + str_i + "\"><v>" + std::to_string(std::get<int64_t>(cell)) +
								 "</v></c>")
									.c_str());
				else if (std::holds_alternative<double>(cell))
					write(("<c r=\"" + str_j + str_i + "\"><v>" + std::to_string(std::get<double>(cell)) +
								 "</v></c>")
									.c_str());
				else
					throw Exception{ "internal error" };
			}
			write("</row>");
		}
		write("</sheetData>");
		write("</worksheet>");
		zip_source_commit_write(source_ptr);

		zip_file_add(zip.archive_ptr_, "xl/worksheets/sheet1.xml", source_ptr, 0);
	}
	if (zip_close(zip.archive_ptr_) < 0)
		throw Exception{ "unable to close the workbook" };
}
} // namespace fd_write_xlsx
#endif // FD_WRITE_XLSX_HPP
