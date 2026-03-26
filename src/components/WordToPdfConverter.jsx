// src/components/WordToPdfConverter.jsx
//import React, { useState, useCallback } from 'react';
import { useDropzone } from 'react-dropzone';
import mammoth from 'mammoth';
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';
import './WordToPdfConverter.css';

const WordToPdfConverter = () => {
  const [file, setFile] = useState(null);
  const [converting, setConverting] = useState(false);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);
  const [preview, setPreview] = useState(null);
  const [showPreview, setShowPreview] = useState(false);
  const [conversionMethod] = useState('styled');
  const [pdfSettings, setPdfSettings] = useState({
    pageSize: 'a4',
    orientation: 'portrait',
    quality: 'high',
    margin: 20,
    preserveFormatting: true
  });
  const [pageRange] = useState({
    start: 1,
    end: 1,
    totalPages: 1
  });

  const onDrop = useCallback(async (acceptedFiles) => {
    const selectedFile = acceptedFiles[0];
    if (selectedFile && (selectedFile.type === 'application/msword' || 
        selectedFile.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')) {
      setFile(selectedFile);
      setError(null);
      setSuccess(null);
      
      try {
        const arrayBuffer = await selectedFile.arrayBuffer();
        
        const options = {
          styleMap: [
            "p[style-name='Title'] => h1.title",
            "p[style-name='Heading1'] => h1.heading1",
            "p[style-name='Heading2'] => h2.heading2",
            "p[style-name='Heading3'] => h3.heading3",
            "p[style-name='Heading4'] => h4.heading4",
            "p[style-name='Heading5'] => h5.heading5",
            "p[style-name='Heading6'] => h6.heading6",
            "r[style-name='Strong'] => strong",
            "r[style-name='Emphasis'] => em"
          ],
          convertImage: mammoth.images.imgElement(function(image) {
            return image.read("base64").then(function(imageBuffer) {
              return {
                src: "data:" + image.contentType + ";base64," + imageBuffer
              };
            });
          })
        };
        
        const result = await mammoth.convertToHtml({ arrayBuffer }, options);
        const styledHtml = `
          <div class="word-content">
            ${result.value}
          </div>
        `;
        
        // Estimar número de páginas baseado no conteúdo
        const textLength = result.value.replace(/<[^>]*>/g, '').length;
        const estimatedPages = Math.max(1, Math.ceil(textLength / 2500));
        
        setPreview(styledHtml);
      } catch (err) {
        console.error('Erro ao gerar preview:', err);
      }
    } else {
      setError('Por favor, selecione um arquivo Word válido (.doc ou .docx)');
      setFile(null);
    }
  }, []);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/msword': ['.doc'],
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx']
    },
    maxFiles: 1
  });

  const convertWithHtml2Canvas = async (htmlContent) => {
    const container = document.createElement('div');
    container.style.position = 'absolute';
    container.style.left = '-9999px';
    container.style.top = '-9999px';
    container.style.width = pdfSettings.orientation === 'portrait' ? '595px' : '842px';
    container.style.backgroundColor = 'white';
    container.style.padding = `${pdfSettings.margin}px`;
    container.innerHTML = htmlContent;
    
    const style = document.createElement('style');
    style.textContent = `
      .word-content {
        font-family: 'Times New Roman', Times, serif;
        line-height: 1.5;
        color: #000;
      }
      .word-content h1, .word-content h2, .word-content h3,
      .word-content h4, .word-content h5, .word-content h6 {
        margin-top: 20px;
        margin-bottom: 10px;
        font-weight: bold;
      }
      .word-content h1 { font-size: 24px; }
      .word-content h2 { font-size: 20px; }
      .word-content h3 { font-size: 18px; }
      .word-content h4 { font-size: 16px; }
      .word-content h5 { font-size: 14px; }
      .word-content h6 { font-size: 12px; }
      .word-content p {
        margin-bottom: 10px;
        font-size: 12px;
      }
      .word-content strong, .word-content b {
        font-weight: bold;
      }
      .word-content em, .word-content i {
        font-style: italic;
      }
      .word-content u {
        text-decoration: underline;
      }
      .word-content ul, .word-content ol {
        margin: 10px 0;
        padding-left: 30px;
      }
      .word-content li {
        margin-bottom: 5px;
      }
      .word-content table {
        border-collapse: collapse;
        width: 100%;
        margin: 10px 0;
      }
      .word-content td, .word-content th {
        border: 1px solid #ddd;
        padding: 8px;
      }
      .word-content th {
        background-color: #f2f2f2;
        font-weight: bold;
      }
      .word-content img {
        max-width: 100%;
        height: auto;
      }
    `;
    container.appendChild(style);
    document.body.appendChild(container);
    
    try {
      const images = container.querySelectorAll('img');
      await Promise.all(Array.from(images).map(img => {
        if (img.complete) return Promise.resolve();
        return new Promise(resolve => {
          img.onload = resolve;
          img.onerror = resolve;
        });
      }));
      
      const canvas = await html2canvas(container, {
        scale: pdfSettings.quality === 'high' ? 3 : pdfSettings.quality === 'normal' ? 2 : 1.5,
        backgroundColor: '#ffffff',
        logging: false,
        useCORS: true
      });
      
      const pdf = new jsPDF({
        orientation: pdfSettings.orientation,
        unit: 'px',
        format: pdfSettings.pageSize
      });
      
      const imgData = canvas.toDataURL('image/png');
      const imgWidth = pdf.internal.pageSize.getWidth();
      const imgHeight = (canvas.height * imgWidth) / canvas.width;
      let heightLeft = imgHeight;
      let position = 0;
      
      pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
      heightLeft -= pdf.internal.pageSize.getHeight();
      
      while (heightLeft > 0) {
        position = heightLeft - imgHeight;
        pdf.addPage();
        pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
        heightLeft -= pdf.internal.pageSize.getHeight();
      }
      
      return pdf;
    } finally {
      document.body.removeChild(container);
    }
  };
  
  const convertWithStyling = async (html) => {
    const pdf = new jsPDF({
      orientation: pdfSettings.orientation,
      unit: 'pt',
      format: pdfSettings.pageSize
    });
    
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = html;
    
    const processElement = (element) => {
      let result = '';
      const tagName = element.tagName?.toLowerCase();
      
      if (element.nodeType === Node.TEXT_NODE) {
        return element.textContent;
      }
      
      if (tagName === 'br') {
        return '\n';
      }
      
      if (tagName === 'strong' || tagName === 'b') {
        const content = Array.from(element.childNodes).map(child => processElement(child)).join('');
        return `**${content}**`;
      }
      
      if (tagName === 'em' || tagName === 'i') {
        const content = Array.from(element.childNodes).map(child => processElement(child)).join('');
        return `*${content}*`;
      }
      
      if (tagName === 'u') {
        const content = Array.from(element.childNodes).map(child => processElement(child)).join('');
        return `_${content}_`;
      }
      
      if (tagName === 'h1' || tagName === 'h2' || tagName === 'h3' || 
          tagName === 'h4' || tagName === 'h5' || tagName === 'h6') {
        const content = Array.from(element.childNodes).map(child => processElement(child)).join('');
        return `\n${content}\n`;
      }
      
      if (tagName === 'p') {
        const content = Array.from(element.childNodes).map(child => processElement(child)).join('');
        return `${content}\n\n`;
      }
      
      if (tagName === 'ul' || tagName === 'ol') {
        const items = Array.from(element.children).map((li, index) => {
          const content = Array.from(li.childNodes).map(child => processElement(child)).join('');
          return tagName === 'ul' ? `• ${content}` : `${index + 1}. ${content}`;
        });
        return items.join('\n') + '\n\n';
      }
      
      return Array.from(element.childNodes).map(child => processElement(child)).join('');
    };
    
    let textContent = processElement(tempDiv);
    const lines = pdf.splitTextToSize(textContent, pdf.internal.pageSize.getWidth() - (pdfSettings.margin * 2));
    
    let yOffset = pdfSettings.margin;
    const pageHeight = pdf.internal.pageSize.getHeight();
    
    for (let i = 0; i < lines.length; i++) {
      let line = lines[i];
      let fontSize = 12;
      let fontWeight = 'normal';
      
      if (line.includes('**')) {
        fontWeight = 'bold';
        line = line.replace(/\*\*/g, '');
      }
      if (line.includes('*')) {
        line = line.replace(/\*/g, '');
      }
      if (line.includes('_')) {
        line = line.replace(/_/g, '');
      }
      
      if (line.trim().startsWith('#')) {
        fontSize = 20;
        fontWeight = 'bold';
        line = line.replace(/#/g, '').trim();
      }
      
      if (yOffset + 15 > pageHeight - pdfSettings.margin) {
        pdf.addPage();
        yOffset = pdfSettings.margin;
      }
      
      pdf.setFontSize(fontSize);
      pdf.setFont('helvetica', fontWeight === 'bold' ? 'bold' : 'normal');
      pdf.text(line, pdfSettings.margin, yOffset);
      yOffset += fontSize + 4;
    }
    
    return pdf;
  };

  const convertToPdf = async () => {
    if (!file) return;

    setConverting(true);
    setError(null);
    setSuccess(null);

    try {
      const arrayBuffer = await file.arrayBuffer();
      
      const options = {
        styleMap: [
          "p[style-name='Title'] => h1.title",
          "p[style-name='Heading1'] => h1.heading1",
          "p[style-name='Heading2'] => h2.heading2",
          "p[style-name='Heading3'] => h3.heading3",
          "p[style-name='Heading4'] => h4.heading4",
          "p[style-name='Heading5'] => h5.heading5",
          "p[style-name='Heading6'] => h6.heading6",
          "r[style-name='Strong'] => strong",
          "r[style-name='Emphasis'] => em",
          "r[style-name='Underline'] => u"
        ],
        convertImage: mammoth.images.imgElement(function(image) {
          return image.read("base64").then(function(imageBuffer) {
            return {
              src: "data:" + image.contentType + ";base64," + imageBuffer
            };
          });
        })
      };
      
      const result = await mammoth.convertToHtml({ arrayBuffer }, options);
      let html = result.value;
      
      html = `
        <div class="word-content" style="font-family: 'Times New Roman', Times, serif; line-height: 1.5;">
          ${html}
        </div>
      `;
      
      let pdf;
      
      if (pdfSettings.preserveFormatting && conversionMethod === 'canvas') {
        pdf = await convertWithHtml2Canvas(html);
      } else {
        pdf = await convertWithStyling(html);
      }
      
      const fileName = file.name.replace(/\.(doc|docx)$/i, '_convertido.pdf');
      pdf.save(fileName);
      
      setSuccess(`Arquivo convertido com sucesso! Download iniciado como ${fileName}`);
      setShowPreview(false);
      setFile(null);
      setPreview(null);
    } catch (err) {
      console.error('Erro na conversão:', err);
      setError('Erro ao converter o arquivo. Por favor, tente novamente.');
    } finally {
      setConverting(false);
    }
  };

  const formatFileSize = (bytes) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  const handleSettingChange = (setting, value) => {
    setPdfSettings(prev => ({
      ...prev,
      [setting]: value
    }));
  };

  return (
    <div className="converter-container">
      <div className="converter-card">
        <h1>Conversor de Word para PDF</h1>
        <p className="subtitle">Converta seus documentos mantendo toda a formatação original</p>
        
        <div 
          {...getRootProps()} 
          className={`dropzone ${isDragActive ? 'active' : ''} ${file ? 'has-file' : ''}`}
        >
          <input {...getInputProps()} />
          <div className="dropzone-content">
            {file ? (
              <div className="file-info">
                <div className="file-icon">📄</div>
                <div className="file-details">
                  <p className="file-name">{file.name}</p>
                  <p className="file-size">{formatFileSize(file.size)}</p>
                </div>
                <button 
                  onClick={(e) => {
                    e.stopPropagation();
                    setFile(null);
                    setError(null);
                    setPreview(null);
                    setShowPreview(false);
                  }} 
                  className="remove-file"
                >
                  ✕
                </button>
              </div>
            ) : (
              <>
                <div className="upload-icon">📤</div>
                <p className="dropzone-text">
                  {isDragActive 
                    ? 'Solte o arquivo aqui...' 
                    : 'Arraste e solte um arquivo Word aqui, ou clique para selecionar'}
                </p>
                <p className="dropzone-subtext">Suporta arquivos .doc e .docx</p>
              </>
            )}
          </div>
        </div>

        {file && (
          <div className="settings-panel">
            <h3>Configurações de Conversão</h3>

            <div className="settings-group">
              <label>Tamanho do Papel:</label>
              <select 
                value={pdfSettings.pageSize} 
                onChange={(e) => handleSettingChange('pageSize', e.target.value)}
                className="setting-select"
              >
                <option value="a4">A4</option>
                <option value="letter">Carta</option>
                <option value="legal">Ofício</option>
                <option value="a3">A3</option>
              </select>
            </div>

            <div className="settings-group">
              <label>Orientação:</label>
              <select 
                value={pdfSettings.orientation} 
                onChange={(e) => handleSettingChange('orientation', e.target.value)}
                className="setting-select"
              >
                <option value="portrait">Retrato</option>
                <option value="landscape">Paisagem</option>
              </select>
            </div>

            <div className="settings-group">
              <label>Qualidade:</label>
              <select 
                value={pdfSettings.quality} 
                onChange={(e) => handleSettingChange('quality', e.target.value)}
                className="setting-select"
              >
                <option value="high">Alta (melhor qualidade)</option>
                <option value="normal">Normal</option>
                <option value="low">Baixa (mais rápido)</option>
              </select>
            </div>

            <div className="settings-group">
              <label>
                <input 
                  type="checkbox" 
                  checked={pdfSettings.preserveFormatting}
                  onChange={(e) => handleSettingChange('preserveFormatting', e.target.checked)}
                />
                Preservar formatação (negrito, itálico, etc.)
              </label>
            </div>

            {preview && (
              <div className="preview-control">
                <button 
                  onClick={() => setShowPreview(!showPreview)}
                  className="preview-button"
                >
                  {showPreview ? 'Ocultar Pré-visualização' : 'Mostrar Pré-visualização'}
                </button>
              </div>
            )}
          </div>
        )}

        {showPreview && preview && (
          <div className="preview-container">
            <h3>Pré-visualização do Documento</h3>
            <div className="preview-content" dangerouslySetInnerHTML={{ __html: preview }} />
          </div>
        )}

        {error && (
          <div className="error-message">
            <span className="error-icon">⚠️</span>
            <p>{error}</p>
          </div>
        )}

        {success && (
          <div className="success-message">
            <span className="success-icon">✅</span>
            <p>{success}</p>
          </div>
        )}

        <button 
          onClick={convertToPdf} 
          disabled={!file || converting}
          className={`convert-button ${(!file || converting) ? 'disabled' : ''}`}
        >
          {converting ? (
            <>
              <span className="spinner"></span>
              Convertendo...
            </>
          ) : (
            'Converter para PDF'
          )}
        </button>

        <div className="features">
          <div className="feature">
            <span className="feature-icon">🎨</span>
            <p>Preserva formatação</p>
          </div>
          <div className="feature">
            <span className="feature-icon">🖼️</span>
            <p>Mantém imagens</p>
          </div>
          <div className="feature">
            <span className="feature-icon">📊</span>
            <p>Suporte a tabelas</p>
          </div>
          <div className="feature">
            <span className="feature-icon">✍️</span>
            <p>Estilos de texto</p>
          </div>
        </div>
      </div>
    </div>
  );
};

export default WordToPdfConverter;