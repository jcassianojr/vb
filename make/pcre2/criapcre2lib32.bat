@echo off
setlocal

echo [1/5] Configurando ambiente MSVC (32-bit)...
call "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\Tools\VsDevCmd.bat" -arch=x86

echo [2/5] Criando diretorio de build...
if not exist build mkdir build
cd build

echo [3/5] Configurando o projeto com CMake (x86)...
:: Aponta para a raiz atual onde os arquivos do PCRE2 estao descompactados
cmake -G "Visual Studio 17 2022" -A Win32 -DCMAKE_INSTALL_PREFIX=C:\devprg\pcre232 ..

echo [4/5] Compilando a Biblioteca (Release)...
cmake --build . --config Release

echo [5/5] Instalando arquivos...
cmake --install . --config Release

cd ..
echo Concluido com sucesso!
pause