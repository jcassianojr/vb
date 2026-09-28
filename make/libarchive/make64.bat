@echo off
setlocal

echo [1/5] Configurando ambiente MSVC (64-bit)...
call "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\Tools\VsDevCmd.bat" -arch=amd64

echo [2/5] Criando diretorio de build 64-bit...
if not exist build64 mkdir build64
cd build64

echo [3/5] Configurando o projeto com CMake...
:: Ajuste o caminho do OpenSSL caso seja diferente de C:\devprg\openssl64
cmake .. -G "NMake Makefiles" ^
    -DCMAKE_BUILD_TYPE=Release ^
    -DCMAKE_INSTALL_PREFIX=C:\devprg\libarchive64 ^
    -DOPENSSL_ROOT_DIR=C:\devprg\openssl64 ^
    -DENABLE_TEST=OFF

echo [4/5] Compilando a biblioteca...
nmake clean
nmake

echo [5/5] Instalando os arquivos...
nmake install

cd ..
echo Concluido com sucesso (64-bit)!
pause