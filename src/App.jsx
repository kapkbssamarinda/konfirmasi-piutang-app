import { useState } from 'react';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';
import './App.css';

// Simple Icon Components
const Icons = {
  Document: () => <span>📄</span>,
  Users: () => <span>👥</span>,
  Download: () => <span>⬇️</span>,
  Upload: () => <span>📤</span>,
  Check: () => <span>✓</span>,
  Info: () => <span>ⓘ</span>,
  Sparkles: () => <span>✨</span>,
  ArrowLeft: () => <span>←</span>,
  File: () => <span>📁</span>,
};

function App() {
  const [templateFile, setTemplateFile] = useState(null);
  const [excelNames, setExcelNames] = useState([]);
  const [manualNames, setManualNames] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [hasGenerated, setHasGenerated] = useState(false);
  const [downloadData, setDownloadData] = useState({ blob: null, fileName: '', isZip: false });
  const [activeStep, setActiveStep] = useState(1);

  const [formData, setFormData] = useState({
    Kota: '', Tanggal_Konfirmasi: '', Periode: '', Nama_Klien: '',
    Sebutan1: '', Auditor1: '', Sebutan2: '', Auditor2: '',
    Tanggal_Jatuh_Tempo: '', Nama_Direktur: '', Jabatan: ''
  });

  const handleInputChange = (e) => {
    const { name, value } = e.target;
    setFormData((prev) => ({ ...prev, [name]: value }));
  };

  const handleExcelUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheet = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheet];
      const json = XLSX.utils.sheet_to_json(worksheet);
      const names = [];
      json.forEach((row) => {
        const nameKey = Object.keys(row).find((key) => key.toLowerCase().includes('nama'));
        if (nameKey && row[nameKey]) names.push(row[nameKey].toString().trim());
      });
      setExcelNames(names);
      setActiveStep(3);
      alert(`${names.length} nama penerima berhasil dimuat dari Excel!`);
    };
    reader.readAsArrayBuffer(file);
  };

  const generateDocuments = async () => {
    if (!templateFile) {
      alert('Harap upload file Template Word terlebih dahulu!');
      return;
    }
    setIsProcessing(true);
    try {
      const manualArray = manualNames.split('\n').map((n) => n.trim()).filter((n) => n);
      const allNames = [...new Set([...excelNames, ...manualArray])];
      if (allNames.length === 0) {
        alert('Harap masukkan setidaknya satu Nama Penerima!');
        setIsProcessing(false);
        return;
      }
      const reader = new FileReader();
      reader.onload = async (event) => {
        try {
          const content = event.target.result;
          const zipResult = new JSZip();
          allNames.forEach((penerima) => {
            const zipTemplate = new PizZip(content);
            const doc = new Docxtemplater(zipTemplate, {
              paragraphLoop: true, linebreaks: true, delimiters: { start: '{{', end: '}}' }
            });
            const docData = { ...formData, Nama_Penerima: penerima };
            Object.keys(docData).forEach(key => {
               if(!docData[key]) docData[key] = `{{${key}}}`; 
               if((key === 'Sebutan1' || key === 'Sebutan2') && docData[key] === `{{${key}}}`) docData[key] = "";
            });
            doc.render(docData);
            const out = doc.getZip().generate({
              type: 'blob',
              mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            });
            if (allNames.length === 1) {
              setDownloadData({
                blob: out,
                fileName: `Konfirmasi Piutang - ${penerima}.docx`,
                isZip: false
              });
            } else {
              zipResult.file(`Konfirmasi Piutang - ${penerima}.docx`, out);
            }
          });
          if (allNames.length > 1) {
            const zipContent = await zipResult.generateAsync({ type: 'blob' });
            setDownloadData({
              blob: zipContent,
              fileName: `Konfirmasi Piutang - ${formData.Nama_Klien || 'Klien'}.zip`,
              isZip: true
            });
          }
          setHasGenerated(true);
        } catch (error) {
          console.error("Error Detail:", error);
          if (error.properties && error.properties.errors instanceof Array) {
            const errorMessages = error.properties.errors.map(err => `- ${err.properties.explanation}`).join("\n");
            alert("Sistem menemukan masalah pada template Anda:\n" + errorMessages);
          } else {
            alert("Terjadi kesalahan: " + error.message);
          }
        } finally {
          setIsProcessing(false);
        }
      };
      reader.readAsArrayBuffer(templateFile);
    } catch (error) {
      console.error(error);
      setIsProcessing(false);
    }
  };

  const totalRecipients = new Set([...excelNames, ...manualNames.split('\n').filter(n => n.trim())].filter(n => n)).size;
  const currentStep = hasGenerated ? 4 : activeStep;

  return (
    <div className="app-wrapper">
      {/* Header */}
      <header className="app-header">
        <div className="app-header__logo">
          <span className="app-header__logo-icon"><Icons.Document /></span>
          <h1 className="app-header__title">Generator Konfirmasi Piutang</h1>
        </div>
        <p className="app-header__subtitle">
          Buat surat konfirmasi piutang untuk audit dengan cepat dan mudah
        </p>
      </header>

      {/* Main Card */}
      <main className="app-card">
        {/* Progress Bar */}
        <div className="progress-bar">
          {[
            { step: 1, label: 'Template' },
            { step: 2, label: 'Detail' },
            { step: 3, label: 'Penerima' },
          ].map(({ step, label }) => (
            <div 
              key={step} 
              className={`progress-bar__step ${step === currentStep ? 'active' : ''} ${step < currentStep ? 'completed' : ''}`}
            >
              <div className="progress-bar__circle">
                {step < currentStep ? <Icons.Check /> : step}
              </div>
              <span className="progress-bar__label">{label}</span>
            </div>
          ))}
        </div>

        {!hasGenerated ? (
          <>
            {/* Step 1: Template */}
            <section className="section">
              <div className="section__header">
                <span className="section__number">1</span>
                <h3 className="section__title">Upload Template</h3>
              </div>
              <p className="section__description">
                Pilih file template Word yang akan digunakan
              </p>
              
              <div className="file-upload mt-3">
                <label className="file-upload__area">
                  <input 
                    type="file" 
                    accept=".docx" 
                    className="file-upload__input"
                    onChange={(e) => {
                      setTemplateFile(e.target.files[0]);
                      setActiveStep(2);
                    }} 
                  />
                  <div className="file-upload__icon"><Icons.Upload /></div>
                  <p className="file-upload__text">
                    <strong>Pilih file</strong> atau drag & drop di sini
                  </p>
                  <p className="form-hint mt-2">Format: .docx • Maks 10MB</p>
                </label>
                
                {templateFile && (
                  <div className="file-upload__preview">
                    <span className="file-upload__preview-icon"><Icons.File /></span>
                    <span>{templateFile.name}</span>
                  </div>
                )}
              </div>
              
              <div className="mt-3">
                <a 
                  href="/bahan/Konfirmasi-Piutang-Template.docx" 
                  download 
                  className="btn btn--outline btn--full"
                >
                  <Icons.Download /> Download Template Standar
                </a>
              </div>
            </section>

            {/* Step 2: Detail Audit */}
            <section className="section">
              <div className="section__header">
                <span className="section__number">2</span>
                <h3 className="section__title">Detail Audit</h3>
              </div>
              <p className="section__description">
                Lengkapi informasi yang akan muncul di semua surat
              </p>
              
              <div className="form-grid mt-3">
                <div className="form-group">
                  <label className="form-label">Kota <span className="form-label__required">*</span></label>
                  <input 
                    name="Kota" 
                    className="form-input"
                    placeholder="Samarinda" 
                    onChange={handleInputChange}
                    value={formData.Kota}
                  />
                </div>
                
                <div className="form-group">
                  <label className="form-label">Tanggal Surat <span className="form-label__required">*</span></label>
                  <input 
                    name="Tanggal_Konfirmasi" 
                    type="date"
                    className="form-input"
                    onChange={handleInputChange}
                    value={formData.Tanggal_Konfirmasi}
                  />
                </div>
                
                <div className="form-group">
                  <label className="form-label">Periode Audit</label>
                  <input 
                    name="Periode" 
                    className="form-input"
                    placeholder="31 Desember 2023" 
                    onChange={handleInputChange}
                    value={formData.Periode}
                  />
                </div>
                
                <div className="form-group">
                  <label className="form-label">Nama Klien <span className="form-label__required">*</span></label>
                  <input 
                    name="Nama_Klien" 
                    className="form-input"
                    placeholder="PT Contoh Abadi" 
                    onChange={handleInputChange}
                    value={formData.Nama_Klien}
                  />
                </div>
                
                <div className="form-row">
                  <div className="form-row__item form-row__item--small">
                    <label className="form-label">Sebutan</label>
                    <input 
                      name="Sebutan1" 
                      className="form-input"
                      placeholder="Bpk" 
                      onChange={handleInputChange}
                      value={formData.Sebutan1}
                    />
                  </div>
                  <div className="form-row__item form-row__item--large">
                    <label className="form-label">Auditor 1</label>
                    <input 
                      name="Auditor1" 
                      className="form-input"
                      placeholder="Nama lengkap" 
                      onChange={handleInputChange}
                      value={formData.Auditor1}
                    />
                  </div>
                </div>
                
                <div className="form-row">
                  <div className="form-row__item form-row__item--small">
                    <label className="form-label">Sebutan</label>
                    <input 
                      name="Sebutan2" 
                      className="form-input"
                      placeholder="Ibu" 
                      onChange={handleInputChange}
                      value={formData.Sebutan2}
                    />
                  </div>
                  <div className="form-row__item form-row__item--large">
                    <label className="form-label">Auditor 2</label>
                    <input 
                      name="Auditor2" 
                      className="form-input"
                      placeholder="Nama lengkap" 
                      onChange={handleInputChange}
                      value={formData.Auditor2}
                    />
                  </div>
                </div>
                
                <div className="form-group">
                  <label className="form-label">Batas Waktu Respon</label>
                  <input 
                    name="Tanggal_Jatuh_Tempo" 
                    type="date"
                    className="form-input"
                    onChange={handleInputChange}
                    value={formData.Tanggal_Jatuh_Tempo}
                  />
                </div>
                
                <div className="form-row">
                  <div className="form-row__item">
                    <label className="form-label">Nama Direktur <span className="form-label__required">*</span></label>
                    <input 
                      name="Nama_Direktur" 
                      className="form-input"
                      placeholder="Nama penanda tangan" 
                      onChange={handleInputChange}
                      value={formData.Nama_Direktur}
                    />
                  </div>
                  <div className="form-row__item" style={{maxWidth: '180px'}}>
                    <label className="form-label">Jabatan</label>
                    <input 
                      name="Jabatan" 
                      className="form-input"
                      placeholder="Direktur" 
                      onChange={handleInputChange}
                      value={formData.Jabatan}
                    />
                  </div>
                </div>
              </div>
            </section>

            {/* Step 3: Daftar Penerima */}
            <section className="section">
              <div className="section__header">
                <span className="section__number">3</span>
                <h3 className="section__title">Daftar Penerima</h3>
              </div>
              <p className="section__description">
                Tambahkan nama penerima konfirmasi (minimal 1)
              </p>
              
              <div className="mt-3">
                <label className="form-label mb-2">Import dari Excel</label>
                <div className="file-upload">
                  <label className="file-upload__area">
                    <input 
                      type="file" 
                      accept=".xlsx, .xls" 
                      className="file-upload__input"
                      onChange={handleExcelUpload} 
                    />
                    <div className="file-upload__icon"><Icons.Users /></div>
                    <p className="file-upload__text">
                      <strong>Upload Excel</strong> dengan kolom "Nama"
                    </p>
                    <p className="form-hint mt-2">Format: .xlsx atau .xls</p>
                  </label>
                  {excelNames.length > 0 && (
                    <div className="file-upload__preview">
                      <Icons.Check /> {excelNames.length} nama berhasil diimport
                    </div>
                  )}
                </div>
              </div>
              
              <div className="divider" />
              
              <div>
                <label className="form-label mb-2">Input Manual</label>
                <textarea
                  className="form-textarea"
                  placeholder="PT Maju Bersama&#10;CV Sejahtera Abadi&#10;Toko Berkah Jaya"
                  value={manualNames}
                  onChange={(e) => {
                    setManualNames(e.target.value);
                    setActiveStep(3);
                  }}
                />
                <p className="form-hint">
                  Total: <strong>{totalRecipients}</strong> penerima
                </p>
              </div>
            </section>

            {/* Action Bar */}
            <div className="action-bar">
              <button 
                className="btn btn--ghost"
                onClick={() => {
                  if (window.confirm('Reset semua data?')) {
                    setTemplateFile(null);
                    setExcelNames([]);
                    setManualNames('');
                    setFormData({
                      Kota: '', Tanggal_Konfirmasi: '', Periode: '', Nama_Klien: '',
                      Sebutan1: '', Auditor1: '', Sebutan2: '', Auditor2: '',
                      Tanggal_Jatuh_Tempo: '', Nama_Direktur: '', Jabatan: ''
                    });
                    setActiveStep(1);
                  }
                }}
              >
                Reset
              </button>
              <button 
                className="btn btn--primary btn--lg" 
                onClick={generateDocuments} 
                disabled={isProcessing || !templateFile}
              >
                {isProcessing ? (
                  <>
                    <span className="btn__spinner"></span>
                    Memproses...
                  </>
                ) : (
                  <>
                    <Icons.Sparkles />
                    Generate {totalRecipients || 1} Dokumen
                  </>
                )}
              </button>
            </div>
          </>
        ) : (
          /* Result Card */
          <section className="section">
            <div className="result-card">
              <div className="result-card__icon"><Icons.Check /></div>
              <span className="result-card__badge"><Icons.Sparkles /> Selesai</span>
              <h4 className="result-card__title">Dokumen Berhasil Dibuat</h4>
              <p className="result-card__message">
                {downloadData.isZip 
                  ? `${totalRecipients} file dikemas dalam format ZIP`
                  : 'File siap diunduh'
                }
              </p>
              
              <div className="result-card__actions">
                <button 
                  className="btn btn--success btn--lg"
                  onClick={() => saveAs(downloadData.blob, downloadData.fileName)}
                >
                  <Icons.Download /> 
                  Unduh {downloadData.isZip ? 'ZIP' : 'File'}
                </button>
                
                <button 
                  className="btn btn--outline"
                  onClick={() => {
                    alert("💡 Tips:\n\n• Buka file di Microsoft Word\n• Gunakan 'Save As' → PDF untuk konversi\n• Format tabel akan tetap rapi");
                  }}
                >
                  <Icons.Info /> Panduan
                </button>
              </div>
            </div>
            
            <div className="action-bar">
              <button 
                className="btn btn--ghost"
                onClick={() => {
                  setHasGenerated(false);
                  setActiveStep(3);
                }}
              >
                <Icons.ArrowLeft /> Edit
              </button>
              <button 
                className="btn btn--outline"
                onClick={() => {
                  setHasGenerated(false);
                  setActiveStep(1);
                  setTemplateFile(null);
                  setExcelNames([]);
                  setManualNames('');
                }}
              >
                <Icons.Sparkles /> Buat Baru
              </button>
            </div>
          </section>
        )}
      </main>

      {/* Footer */}
      <footer className="text-center text-muted" style={{ fontSize: '0.8125rem' }}>
        <p>Generator Konfirmasi Piutang</p>
      </footer>
    </div>
  );
}

export default App;