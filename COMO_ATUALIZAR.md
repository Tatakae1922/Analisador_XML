# Como atualizar o Analisador de NF-e HEC e gerar um novo executável

Este documento explica o processo usado neste projeto para publicar uma
nova versão do programa (código-fonte + executável), para qualquer
sessão/chat que for mexer nele depois consiga seguir o mesmo padrão.

Repositório: https://github.com/Tatakae1922/Analisador_XML

## 1. Estrutura do projeto

- `analisador_xml.py` — código-fonte único do programa (GUI + lógica de
  extração dos XMLs). Tudo fica neste arquivo, sem módulos separados.
- `logo.png` / `logo_hec.ico` — logo da HEC usada na interface e como
  ícone do executável.
- `Atualizar.bat` — script que o **usuário final** roda (não o
  desenvolvedor). Sincroniza a pasta do programa a partir de um
  servidor de rede para uma pasta local e abre a cópia local. Ver
  seção 4.
- `servidor.txt` / `servidor.txt.exemplo` — `servidor.txt` guarda o
  caminho real do servidor (ex.: `\\prosrv\...\ANALISADOR_XML`) e
  **não é versionado** (está no `.gitignore`) porque o repositório é
  público e esse caminho é informação interna da rede. O
  `.exemplo` documenta o formato esperado com um caminho fictício.
- `.gitignore` — ignora `build/`, `dist/`, `*.spec`, `tema_preferido.txt`
  e `servidor.txt` (artefatos locais/de build, não fazem parte do
  código-fonte).

## 2. Convenção de versão

A versão do programa fica numa única constante, perto do topo do
arquivo:

```python
VERSAO_PROGRAMA = "v01.6"  # atualize a cada nova versao gerada
```

**Sempre que o código for alterado e for gerar um executável novo,
incremente essa constante primeiro** (ex.: `v01.6` -> `v01.7`). Ela
aparece no rodapé do menu lateral e no título da janela, então é a
única fonte de verdade sobre "qual versão é essa" para quem for
conferir na tela.

## 3. Passo a passo para publicar uma atualização

1. **Editar `analisador_xml.py`** com a mudança desejada.
2. **Dar bump em `VERSAO_PROGRAMA`** (seção 2 acima).
3. **Compilar/testar antes de empacotar**:
   ```bash
   python -m py_compile analisador_xml.py
   ```
   Depois, rodar o `.py` direto (`python analisador_xml.py`) e conferir
   que a janela abre sem erro — encerrar o processo depois de
   confirmar (`taskkill /F /IM python.exe` no Windows).
4. **Commitar e subir pro GitHub** (`git add`, `git commit`, `git
   push` no repositório acima). Isso deve acontecer **antes** de gerar
   o executável, para o histórico do código bater com a versão do
   `.exe`.
5. **Gerar o executável com PyInstaller** (modo `--onedir`, não
   `--onefile` — onedir abre mais rápido, e é o padrão dos outros
   programas HEC):
   ```bash
   python -m PyInstaller --onedir --windowed \
     --name "Analisador_NFe_HEC_vXX.Y" \
     --icon logo_hec.ico \
     --add-data "logo.png;." --add-data "logo_hec.ico;." \
     analisador_xml.py
   ```
   Troque `vXX.Y` pela versão atual (mesma do passo 2). O resultado
   fica em `dist/Analisador_NFe_HEC_vXX.Y/` (o `.exe` + uma pasta
   `_internal` — **os dois são necessários**, não dá pra distribuir só
   o `.exe` sozinho).
6. **Testar o executável gerado** (não só o `.py`): rodar o `.exe` de
   dentro de `dist/...`, confirmar que abre sem erro, e opcionalmente
   conferir que o ícone da HEC aparece certo na barra de título/ícone
   do arquivo.
7. **Compactar a pasta `dist/Analisador_NFe_HEC_vXX.Y` inteira em um
   `.zip`** antes de distribuir (o usuário final precisa extrair a
   pasta inteira, não só copiar o `.exe`).
8. **Limpar arquivos temporários locais** antes de seguir (não
   versionados, mas podem sobrar do teste): `tema_preferido.txt`,
   `servidor.txt`, `__pycache__/`. `build/`, `dist/` e `*.spec` já
   ficam de fora do git pelo `.gitignore`.

## 4. Como o usuário final recebe as atualizações (`Atualizar.bat`)

O `Atualizar.bat` **não é rodado por quem desenvolve** o programa — é
o atalho que o usuário final (contador, equipe) usa no dia a dia:

1. Na primeira execução em um computador, pergunta o caminho da pasta
   do servidor (rede interna) e salva em `servidor.txt` (local, ao
   lado do `.bat`, nunca no Git).
2. Nas execuções seguintes, lê `servidor.txt` direto.
3. Usa `robocopy /MIR` para sincronizar a pasta do servidor com uma
   pasta local (`%LOCALAPPDATA%\AnalisadorNFeHEC`) — só baixa o que
   mudou, rápido depois da primeira vez.
4. Procura um arquivo `Analisador_NFe_HEC_*.exe` dentro da pasta local
   sincronizada e abre ele.

**Importante:** para uma atualização chegar aos usuários finais, além
de gerar o novo `.exe` (seção 3), é preciso **colocar a pasta
`Analisador_NFe_HEC_vXX.Y` (extraída do zip) dentro da pasta do
servidor** que está configurada no `servidor.txt` de cada usuário — o
`Atualizar.bat` sincroniza dessa pasta do servidor, não do GitHub.

## 5. Resumo rápido (checklist)

- [ ] Código alterado em `analisador_xml.py`
- [ ] `VERSAO_PROGRAMA` incrementado
- [ ] `python -m py_compile` sem erro + teste manual do `.py`
- [ ] `git add` + `git commit` + `git push`
- [ ] `pyinstaller --onedir --windowed --name "Analisador_NFe_HEC_vXX.Y" ...`
- [ ] Testar o `.exe` gerado (abre sem erro, ícone certo)
- [ ] Compactar `dist/Analisador_NFe_HEC_vXX.Y` em `.zip` e entregar
- [ ] Colocar a pasta nova no servidor de rede (pra quem usa `Atualizar.bat`)
