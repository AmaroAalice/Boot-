# WhatsApp Message Automation

[![Python](https://img.shields.io/badge/python-3.10+-blue)](https://www.python.org/)
[![Selenium](https://img.shields.io/badge/selenium-4.15.0-orange)](https://selenium.dev/)

Automação de envio de mensagens personalizadas para clientes via WhatsApp Web utilizando **Python** e **Selenium**. O script lê uma planilha Excel com informações de clientes e envia mensagens diretamente para os contatos listados.

---

## 📋 Funcionalidades

- Lê planilhas `.xlsx` com dados de clientes.
- Gera mensagens personalizadas baseadas em datas e motivos.
- Abre conversas no WhatsApp Web usando Selenium.
- Envia a mensagem automaticamente.
- Registra erros em um arquivo `erros.csv` caso não consiga enviar a mensagem.

---

## ⚙️ Requisitos

- Python 3.10+
- Google Chrome ou Chromium instalado
- ChromeDriver compatível com a versão do Chrome/Chromium
- Bibliotecas Python:

```bash
pip install selenium openpyxl
```
