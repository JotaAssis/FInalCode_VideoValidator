# **YOLOv8 Video Analyzer**

Este projeto usa **YOLOv8** para analisar vídeos a partir de URLs fornecidas em uma planilha Excel. O objetivo é detectar **celulares, smartphones e rádios** nos vídeos e gerar uma nova planilha com os resultados.

---

## **📌 Funcionalidades**

✅ Lê uma planilha Excel com **IDs e URLs de vídeos**.\
✅ Analisa os vídeos e verifica se há **celulares, smartphones ou rádios**.\
✅ Salva os resultados em uma nova planilha com as colunas:

- **ID**
- **Links**
- **Status** → "Pendente", "Validado" ou "Erro"
- **Evidencia** → "Verdadeiro" (se encontrou) ou "Falso" (se não encontrou)

---

## **📂 Estrutura do Projeto**

```
📎 iaGeradaGPT
│── 📂 models
│   │── yolov8n.pt         # Modelo YOLOv8 pré-treinado
│── 📂 data
│   │── videos.xlsx        # Planilha de entrada com IDs e links dos vídeos
│   │── resultado.xlsx     # Planilha de saída com os resultados
│── main.py                # Código principal do projeto
│── requirements.txt       # Dependências do projeto
│── README.md              # Documentação do projeto
```

---

## **🚀 Como Configurar o Projeto**

### **1 Criar um ambiente virtual (opcional, mas recomendado)**

```bash
python -m venv venv
source venv/bin/activate  # Linux/macOS
venv\Scripts\activate     # Windows
```

### **2 Instalar as dependências**

```bash
pip install -r requirements.txt
```

Caso precise instalar manualmente:

```bash
pip install ultralytics pandas opencv-python openpyxl
```

### **3 Baixar o modelo YOLOv8**

O modelo será baixado automaticamente ao rodar o código.\
Se precisar baixar manualmente, execute:

```python
from ultralytics import YOLO
YOLO("yolov8n.pt")  # Baixa o modelo YOLOv8 Nano
```

Mova o arquivo para a pasta `models/`.

---

## **▶️ Como Executar**

1. **Edite a planilha ****`data/videos.xlsx`** e adicione os vídeos a serem analisados.
2. **Execute o script principal**:
   ```bash
   python main.py
   ```
3. O resultado será salvo em `data/resultado.xlsx`.

---

## **📊 Exemplo de Entrada (****`videos.xlsx`****)**

| ID | Links                                                            |
| -- | ---------------------------------------------------------------- |
| 1  | [https://example.com/video1.mp4](https://example.com/video1.mp4) |
| 2  | [https://example.com/video2.mp4](https://example.com/video2.mp4) |

---

## **📊 Exemplo de Saída (****`resultado.xlsx`****)**

| ID | Links                                                            | Status   | Evidencia  |
| -- | ---------------------------------------------------------------- | -------- | ---------- |
| 1  | [https://example.com/video1.mp4](https://example.com/video1.mp4) | Validado | Verdadeiro |
| 2  | [https://example.com/video2.mp4](https://example.com/video2.mp4) | Validado | Falso      |

---

## **🛠 Possíveis Erros e Soluções**

**Erro:** `ModuleNotFoundError: No module named 'ultralytics'`\
✅ Solução: Execute `pip install ultralytics`

**Erro:** `No such file or directory: 'models/yolov8n.pt'`\
✅ Solução: Baixe o modelo manualmente ou execute `python main.py` para que ele seja baixado automaticamente.

---

## **📌 Melhorias Futuras**

🔹 Suporte para mais tipos de objetos.\
🔹 Melhor otimização para processar vídeos mais rapidamente.\
🔹 Interface gráfica para facilitar o uso.

---

### **👨‍💻 Autor**

📌 Desenvolvido por João Paulo

