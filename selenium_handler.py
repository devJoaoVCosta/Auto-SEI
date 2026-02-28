from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import os
from datetime import datetime
import logging
import traceback
import pyautogui
import time


class SEIAutomation:
    def __init__(self):
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()
        self.wait = WebDriverWait(self.driver, 10)
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)

    def executar(self, usuario, senha, processo, documentos: list) -> list:
        """
        Retorna uma lista de booleanos indicando sucesso/falha por documento.
        documentos: lista de dicts com chaves 'tipo' e 'caminho'.
        """
        resultados = []
        try:
            self.logger.info("Iniciando automacao")
            self.login(usuario, senha)
            self.buscar_processo(processo)
            self.logger.info(f"Processo {processo} aberto. Processando {len(documentos)} documento(s)...")

            for i, doc in enumerate(documentos, 1):
                tipo = doc["tipo"]
                caminho = doc["caminho"]
                self.logger.info(f"[{i}/{len(documentos)}] Tipo: {tipo} | Arquivo: {os.path.basename(caminho)}")
                try:
                    self.incluir_documento(tipo, caminho)
                    resultados.append(True)
                    self.logger.info(f"[{i}/{len(documentos)}] Sucesso.")
                except Exception as e:
                    self.logger.error(f"[{i}/{len(documentos)}] Falha: {str(e)}")
                    self.logger.error(traceback.format_exc())
                    resultados.append(False)
                    try:
                        self.driver.switch_to.default_content()
                    except:
                        pass

        except Exception as e:
            self.logger.error(f"Erro geral: {str(e)}")
            self.logger.error(traceback.format_exc())
            raise
        finally:
            self.logger.info("Encerrando automacao")
            try:
                self.driver.quit()
            except:
                pass

        return resultados

    def login(self, usuario, senha):
        self.driver.get("https://sei.funprespjud.com.br/")
        self.wait.until(
            EC.presence_of_element_located((By.ID, "txtUsuario"))
        ).send_keys(usuario)
        self.driver.find_element(By.ID, "pwdSenha").send_keys(senha)
        self.driver.find_element(By.ID, "sbmAcessar").click()

        # Aguarda 3s e verifica se o SEI abriu alerta de credenciais invalidas
        time.sleep(3)
        try:
            alert = self.driver.switch_to.alert
            mensagem = alert.text
            alert.accept()
            raise Exception(f"Erro de login: {mensagem}")
        except Exception as e:
            if "Erro de login" in str(e):
                raise
            # Nao havia alerta — login bem-sucedido
            self.logger.info("Login realizado com sucesso")

    def buscar_processo(self, processo):
        campo = self.wait.until(
            EC.presence_of_element_located((By.ID, "txtPesquisaRapida"))
        )
        campo.clear()
        campo.send_keys(processo)
        campo.send_keys(Keys.RETURN)

    def escrever_texto_robusto(self, texto, intervalo=0.025, tentativas=3):
        """
        Cola o texto via area de transferencia (Ctrl+V) para suportar
        caracteres especiais como accentos, cedilha, chaves, etc.
        O pyautogui.write() nao consegue digitar esses caracteres corretamente.
        """
        import tkinter as tk

        # Usa tkinter para copiar para a area de transferencia sem dependencia extra
        root = tk.Tk()
        root.withdraw()
        root.clipboard_clear()
        root.clipboard_append(texto)
        root.update()

        time.sleep(0.3)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.5)
        self.logger.info(f"Texto colado via Ctrl+V: {texto}")

        root.destroy()

    def incluir_documento(self, tipo_documento: str, caminho_arquivo: str):
        """
        Cria um documento externo no processo aberto.
        tipo_documento: texto exato que aparece no <select> do SEI (ex: 'Comprovante').
        caminho_arquivo: caminho absoluto do PDF a ser anexado.
        """
        self.logger.info(f"Incluindo '{tipo_documento}' | {caminho_arquivo}")
        self.driver.implicitly_wait(5)

        # ────────────────────────────────────────────────
        # PARTE 1 — Clicar em "Incluir Documento"
        # ────────────────────────────────────────────────
        for tentativa in range(3):
            try:
                self.logger.info(f"Tentativa {tentativa + 1}: clicar em 'Incluir Documento'")
                try:
                    frame = self.wait.until(
                        EC.presence_of_element_located((By.NAME, "ifrConteudoVisualizacao"))
                    )
                    self.driver.switch_to.frame(frame)
                except:
                    pass

                clicked = self.driver.execute_script("""
                    var links = document.getElementsByTagName('a');
                    for (var i = 0; i < links.length; i++) {
                        var l = links[i];
                        if (l.href.includes('acao=documento_escolher_tipo&acao_origem') &&
                            l.querySelector('img[src*="documento_incluir.svg?18"]')) {
                            l.click(); return true;
                        }
                    }
                    return false;
                """)

                if clicked:
                    self.logger.info("'Incluir Documento' clicado")

                    # PARTE 2 — Selecionar "Externo"
                    try:
                        self.wait.until(
                            EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrVisualizacao"))
                        )
                    except:
                        pass

                    externo_ok = False
                    for t in range(3):
                        try:
                            externo_ok = self.driver.execute_script("""
                                var links = document.getElementsByTagName('a');
                                for (var i = 0; i < links.length; i++) {
                                    var l = links[i];
                                    if (l.href.includes('acao=documento_escolher_tipo&acao_origem') &&
                                        l.textContent.trim() === 'Externo' &&
                                        l.className === 'ancoraOpcao') {
                                        l.click(); return true;
                                    }
                                }
                                return false;
                            """)
                            if externo_ok:
                                self.logger.info("'Externo' clicado")
                                break
                        except:
                            pass
                        time.sleep(2)

                    try:
                        self.driver.switch_to.default_content()
                    except:
                        pass

                    if externo_ok:
                        break
                    else:
                        self.logger.warning("'Externo' nao encontrado, repetindo...")
                        time.sleep(1)
                        continue

            except Exception as e:
                self.logger.error(f"Erro tentativa {tentativa + 1}: {e}")
                if tentativa == 2:
                    raise
            try:
                self.driver.switch_to.default_content()
            except:
                pass
            time.sleep(2)

        # ────────────────────────────────────────────────
        # PARTE 3 — Preencher formulario
        # ────────────────────────────────────────────────
        self.logger.info("Preenchendo formulario")
        self.driver.switch_to.default_content()
        time.sleep(2)

        try:
            self.wait.until(
                EC.frame_to_be_available_and_switch_to_it((By.NAME, "ifrConteudoVisualizacao"))
            )
        except:
            pass

        try:
            self.wait.until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrVisualizacao"))
            )
        except:
            raise Exception("Iframe do formulario nao localizado")

        try:
            # Tipo de Documento — usa o valor dinamico recebido como parametro
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "selSerie"))
            )

            sucesso_tipo = self.driver.execute_script(
                """
                var tipo = arguments[0];
                var sel = document.getElementById('selSerie');
                if (!sel) return false;
                for (var i = 0; i < sel.options.length; i++) {
                    if (sel.options[i].text === tipo) {
                        sel.selectedIndex = i;
                        sel.value = sel.options[i].value;
                        sel.dispatchEvent(new Event('change', { bubbles: true }));
                        sel.dispatchEvent(new Event('input',  { bubbles: true }));
                        sel.dispatchEvent(new Event('blur',   { bubbles: true }));
                        return true;
                    }
                }
                return false;
                """,
                tipo_documento,
            )

            if not sucesso_tipo:
                raise Exception(
                    f"Tipo de documento '{tipo_documento}' nao encontrado no SEI. "
                    f"Verifique se o nome esta correto e disponivel para este processo."
                )

            valor_selecionado = self.driver.execute_script(
                "var s = document.getElementById('selSerie');"
                "return s ? s.options[s.selectedIndex].text : '';"
            )
            if valor_selecionado != tipo_documento:
                raise Exception(
                    f"Selecao incorreta: esperado '{tipo_documento}', obtido '{valor_selecionado}'"
                )

            self.logger.info(f"Tipo de Documento '{tipo_documento}' selecionado")
            time.sleep(2)

            # Data atual
            data_atual = datetime.now().strftime("%d/%m/%Y")
            self.driver.execute_script(
                f"document.getElementById('txtDataElaboracao').value = '{data_atual}';"
            )
            self.logger.info(f"Data: {data_atual}")

            # Nato-digital
            self.driver.execute_script("""
                var r = document.getElementById('optNato');
                if (r) { r.checked = true; r.click(); }
            """)
            time.sleep(1)

            # Nivel de acesso
            nivel = self.driver.execute_script("""
                var pub = document.getElementById('optPublico');
                var res = document.getElementById('optRestrito');
                if (res && res.checked) return 'Restrito (ja selecionado)';
                if (pub && pub.checked) return 'Publico (ja selecionado)';
                if (pub && !pub.disabled) {
                    pub.checked = true; pub.click();
                    pub.dispatchEvent(new Event('change', { bubbles: true }));
                    return 'Publico';
                }
                if (res && !res.disabled) {
                    res.checked = true; res.click();
                    res.dispatchEvent(new Event('change', { bubbles: true }));
                    return 'Restrito';
                }
                return false;
            """)

            if not nivel:
                raise Exception("Nenhum nivel de acesso disponivel")
            self.logger.info(f"Nivel de acesso: {nivel}")
            time.sleep(1)

            # Verificacoes finais
            nato = self.driver.execute_script('return document.getElementById("optNato").checked;')
            pub  = self.driver.execute_script('return document.getElementById("optPublico").checked;')
            res  = self.driver.execute_script('return document.getElementById("optRestrito").checked;')

            if not nato:
                raise Exception("Nato-digital nao foi selecionado")
            if not pub and not res:
                raise Exception("Nivel de acesso nao foi selecionado")

            self.logger.info(f"Formulario OK: Nato={nato}, Publico={pub}, Restrito={res}")

        except Exception as e:
            self.logger.error(f"Erro no formulario: {e}")
            raise

        # ────────────────────────────────────────────────
        # PARTE 4 — Anexar arquivo e Salvar
        # Usa send_keys direto no <input type="file"> oculto,
        # SEM tornar o elemento visivel (evita abertura do
        # explorador de arquivos nativo do Windows).
        # ────────────────────────────────────────────────
        self.logger.info("Anexando arquivo...")

        self.wait.until(EC.presence_of_element_located((By.ID, "frmAnexos")))

        # Normaliza o caminho para barras invertidas (padrao Windows)
        # e garante que caracteres especiais (acentos, etc.) sejam preservados
        caminho_arquivo = os.path.normpath(caminho_arquivo)
        self.logger.info(f"Caminho do arquivo: {caminho_arquivo}")

        # Selenium interage com input[type=file] oculto diretamente,
        # sem precisar alterar CSS ou tornar o elemento visivel.
        # Alterar o CSS pode disparar eventos JS do SEI que abrem
        # o dialogo nativo do Windows, travando a automacao.
        try:
            file_input = self.wait.until(
                EC.presence_of_element_located((By.ID, "inputFile"))
            )
        except:
            file_input = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
            )

        file_input.send_keys(caminho_arquivo)
        self.logger.info("Arquivo anexado com sucesso")
        time.sleep(2)

        # Clica no botao Salvar direto pelo DOM — sem pyautogui
        salvo = self.driver.execute_script("""
            // Tenta pelos IDs mais comuns do SEI
            var ids = ['btnSalvar', 'sbmSalvar', 'btnConfirmar'];
            for (var i = 0; i < ids.length; i++) {
                var el = document.getElementById(ids[i]);
                if (el) { el.click(); return 'id:' + ids[i]; }
            }
            // Fallback: procura por botao/input com texto "Salvar" ou "Confirmar"
            var todos = document.querySelectorAll('button, input[type="submit"], input[type="button"], a');
            for (var j = 0; j < todos.length; j++) {
                var txt = (todos[j].value || todos[j].textContent || '').trim().toLowerCase();
                if (txt === 'salvar' || txt === 'confirmar') {
                    todos[j].click();
                    return 'text:' + txt;
                }
            }
            return null;
        """)

        if salvo:
            self.logger.info(f"Botao Salvar clicado via DOM ({salvo})")
        else:
            # Fallback com pyautogui caso o botao nao seja localizado no DOM
            self.logger.warning("Botao Salvar nao localizado no DOM — usando pyautogui como fallback")
            pyautogui.press("tab")
            time.sleep(0.5)
            pyautogui.press("enter")

        time.sleep(3)
        self.logger.info(f"Documento '{os.path.basename(caminho_arquivo)}' salvo no SEI")
        self.driver.switch_to.default_content()