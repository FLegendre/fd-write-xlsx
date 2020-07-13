all : test format 

test : test.cpp fd-write-xlsx.hpp
	g++ -std=c++17 -Wall -g test.cpp -lzip --output test

format :
	clang-format -i fd-write-xlsx.hpp test.cpp
