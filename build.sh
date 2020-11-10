#!/bin/bash
g++ -static -o olecontainer -mwindows -DUNICODE -D_UNICODE CContainer.cpp CSite.cpp -lole32 -luuid -loledlg
