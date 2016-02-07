all: main.cpp
	clang++ -O3 --std=c++11 -I. -I/usr/local/include -I/Users/baidu/workspace/commercial_software/libxl-mac-3.6.5.2/include_cpp -L. -L/usr/local/lib -lxl -Wl,-rpath,/Users/baidu/workspace/dashifu -lboost_system -lboost_filesystem main.cpp -lgflags -o dashifu 
