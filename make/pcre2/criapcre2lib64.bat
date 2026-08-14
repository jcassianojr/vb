@echo off
setlocal

echo [1/5] Configurando ambiente MSVC (64-bit)...
call "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\Tools\VsDevCmd.bat" -arch=amd64

echo [2/5] Criando diretorio de build (x64)...
if not exist build64 mkdir build64
cd build64

echo [3/5] Configurando o projeto com CMake para biblioteca estatica (x64)...
cmake -G "Visual Studio 17 2022" -A x64 -DCMAKE_INSTALL_PREFIX=C:\devprg\pcre264 ..

echo [4/5] Compilando a Biblioteca (Release x64)...
cmake --build . --config Release

echo [5/5] Instalando arquivos...
cmake --install . --config Release

cd ..
echo Concluido com sucesso!
pause