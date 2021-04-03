all : test-header-only fd-write-xlsx.a test format 

test-header-only : fd-write-xlsx-header-only.hpp test-header-only.cpp
	g++ -std=c++17 -Wall -g test-header-only.cpp -lzip --output test-header-only

# Make a static library.
fd-write-xlsx.a : fd-write-xlsx.cpp
	g++ -std=c++17 -Wall -o9 -c fd-write-xlsx.cpp --output fd-write-xlsx.a

# Test with static library.
test : fd-write-xlsx.hpp test.cpp
	g++ -std=c++17 -Wall -g test.cpp fd-write-xlsx.a -lzip --output test

format :
	clang-format -i fd-write-xlsx-header-only.hpp fd-write-xlsx.hpp fd-write-xlsx.cpp test.cpp test-header-only.cpp
