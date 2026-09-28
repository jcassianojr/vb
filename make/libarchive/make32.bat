@echo off
setlocal

echo [1/5] Configurando ambiente MSVC (32-bit)...
call "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\Tools\VsDevCmd.bat" -arch=x86

echo [2/5] Criando diretorio de build 32-bit...
if not exist build32 mkdir build32
cd build32

echo [3/5] Configurando o projeto com CMake...
:: Ajuste o caminho do OpenSSL caso seja diferente de C:\devprg\openssl32
cmake .. -G "NMake Makefiles" ^
    -DCMAKE_BUILD_TYPE=Release ^
    -DCMAKE_INSTALL_PREFIX=C:\devprg\libarchive32 ^
    -DOPENSSL_ROOT_DIR=C:\devprg\openssl32 ^
    -DENABLE_TEST=OFF

echo [4/5] Compilando a biblioteca...
nmake clean
nmake

echo [5/5] Instalando os arquivos...
nmake install

cd ..
echo Concluido com sucesso (32-bit)!
pause