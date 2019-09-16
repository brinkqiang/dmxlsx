
// Copyright (c) 2018 brinkqiang (brink.qiang@gmail.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

#include "dmxlsx_module.h"

Cdmxlsx_module::Cdmxlsx_module()
{

}

Cdmxlsx_module::~Cdmxlsx_module()
{

}

void DMAPI Cdmxlsx_module::Release(void)
{
    delete this;
}

void DMAPI Cdmxlsx_module::Test(void)
{
    std::cout << "PROJECT_NAME = dmxlsx" << std::endl;
    std::cout << "PROJECT_NAME_UP = DMXLSX" << std::endl;
    std::cout << "PROJECT_NAME_LO = dmxlsx" << std::endl;
}

Idmxlsx* DMAPI dmxlsxGetModule()
{
    return new Cdmxlsx_module();
}
