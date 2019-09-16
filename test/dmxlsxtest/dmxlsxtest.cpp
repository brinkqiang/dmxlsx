
#include "dmxlsx.h"

int main( int argc, char* argv[] ) {

    Idmxlsx* module = dmxlsxGetModule();
    if (module)
    {
        module->Test();
        module->Release();
    }
    return 0;
}
