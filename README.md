# OAB-process-collector
[![NPM](http://img.shields.io/npm/l/react)](https://github.com/fonteneleantp/OAB-process-collector/blob/main/LICENSE)

## Tutorial interface
![OAB process collector tutorial](https://github.com/fonteneleantp/OAB-process-collector/assets/140100514/f237eb50-3009-4c7b-8a3d-4d89d454ed68)

## Sobre o projeto
Projeto desenvolvido afim automatizar o processo de busca e coleta de dados de processos vinculados ao número OAB de advogados específicos. O programa busca dos dados por estado do Brasil e os registra em sheets numa planilha de excel.
Os dados são públicos e disponibilizados em "https://pje-consulta-publica.tjmg.jus.br/"

## Funcionalidade  
O usuário deve indicar o número do registro OAB e selecionar os estados do Brasil em que os processos devem ser buscados, ao clicar em "FIND" o programa irá, de modo automatizado, entrar no site PJE e buscar pelos processos vinculados ao registro OAB naquele estado, após isso ele irá entrar em cada um dos processos e coletar os dados de movimentações, em seguida os dados são registrados numa planilha de excel conforme destacado abaixo para que sejam trabalhados conforme o usuário desejar.
![image](https://github.com/fonteneleantp/OAB-process-collector/assets/140100514/43500719-7e7c-4544-8735-66bfa27c2e30)


## Pré-requesitos
- Google Chrome
- Excel/LibreOffice Calc
- Ter uma planilha excel chamada "dados" no mesmo diretório
- Acesso a internet

## Tecnologia usada
- Python

## Bilbiotecas
- Selenium
- CustomTkinter
- Pillow
- Openpyxl

## Autor
Antonio Pereira Fontenele  
https://www.linkedin.com/in/antonio-fontenele-7555b7179/
