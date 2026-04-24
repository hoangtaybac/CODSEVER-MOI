# MathType Converter API

Chuyển đổi phương trình MathType từ file Word (.docx) sang MathML và LaTeX.

**Sử dụng Ruby gem `mathtype_to_mathml`** - giải pháp đã được chứng minh hoạt động tốt.

## 🚀 Deploy lên Railway

1. Push code lên GitHub
2. Vào [Railway](https://railway.app) → New Project → Deploy from GitHub
3. Chọn repo này
4. Đợi build (~2-3 phút do cần cài Ruby gem)
5. Generate domain trong Settings → Networking

## 📁 Cấu trúc

```
mathtype-converter/
├── Dockerfile          ← Node.js + Ruby
├── package.json        
├── server.js           ← Express API
├── mt2mml.rb          ← Gọi gem mathtype_to_mathml
└── README.md
```

## 🔧 API Endpoints

### `GET /health`
```bash
curl https://your-app.railway.app/health
```

### `POST /convert`
Convert single OLE file (.bin)
```bash
curl -X POST https://your-app.railway.app/convert \
  -F "file=@oleObject1.bin"
```

### `POST /convert-docx`
Convert all equations from .docx
```bash
curl -X POST https://your-app.railway.app/convert-docx \
  -F "file=@document.docx"
```

## 📝 Response Example

```json
{
  "success": true,
  "total": 44,
  "errors": 0,
  "equations": [
    {
      "index": 1,
      "name": "oleObject1.bin",
      "mathml": "<math xmlns='...'><mfrac>...</mfrac></math>",
      "latex": "\\frac{a}{b}",
      "error": null
    }
  ]
}
```

## 🖥️ Chạy Local

```bash
# Cần cài Ruby và gem trước
gem install mathtype_to_mathml

# Install dependencies
npm install

# Run
npm start
# hoặc
node server.js
```

## 🐳 Docker Local

```bash
docker build -t mathtype-converter .
docker run -p 8000:8000 mathtype-converter
```

## ⚙️ Tech Stack

- **Node.js 20** - Express server
- **Ruby** - gem `mathtype_to_mathml` để parse MTEF
- **mathml-to-latex** - npm package để convert MathML → LaTeX
