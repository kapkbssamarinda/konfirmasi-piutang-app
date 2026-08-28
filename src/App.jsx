import { useState, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { useAutoAnimate } from '@formkit/auto-animate/react';
import ReactConfetti from 'react-confetti';
import { useWindowSize } from 'react-use';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';
import logoTransparan from './assets/logo_transparan.png';
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
  Plus: () => <span>➕</span>,
  Trash: () => <span>🗑️</span>,
  Table: () => <span>📋</span>,
  Money: () => <span>💰</span>,
  Edit: () => <span>✏️</span>,
  Calendar: () => <span>📅</span>,
};

// Indonesian Month Names
const BULAN_INDONESIA = [
  'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
];

// Helper: Format any date (Date object, Excel serial number, ISO/UK date string) to "1 Januari 2025"
const formatIndonesianDate = (val) => {
  if (!val && val !== 0) return '';

  // 1. If it's a JS Date object
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return `${val.getDate()} ${BULAN_INDONESIA[val.getMonth()]} ${val.getFullYear()}`;
  }

  // 2. If it's an Excel numeric serial number (e.g. 45658 for 2025-01-01)
  if (typeof val === 'number') {
    if (val > 1000 && val < 100000) {
      if (typeof XLSX !== 'undefined' && XLSX.SSF && XLSX.SSF.parse_date_code) {
        const parsed = XLSX.SSF.parse_date_code(val);
        if (parsed && parsed.y && parsed.m && parsed.d) {
          return `${parsed.d} ${BULAN_INDONESIA[parsed.m - 1]} ${parsed.y}`;
        }
      }
      const date = new Date(Math.round((val - 25569) * 86400 * 1000));
      if (!isNaN(date.getTime())) {
        return `${date.getDate()} ${BULAN_INDONESIA[date.getMonth()]} ${date.getFullYear()}`;
      }
    }
  }

  const str = val.toString().trim();
  if (!str) return '';

  // 3. If already formatted in Indonesian (e.g. "1 Januari 2025")
  const hasIndonesianMonth = BULAN_INDONESIA.some((m) => new RegExp(`\\b${m}\\b`, 'i').test(str));
  if (hasIndonesianMonth) return str;

  // 4. ISO Date format: YYYY-MM-DD or YYYY/MM/DD (e.g. 2025-01-01 or 2025-1-1)
  const isoMatch = str.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (isoMatch) {
    const year = parseInt(isoMatch[1], 10);
    const month = parseInt(isoMatch[2], 10);
    const day = parseInt(isoMatch[3], 10);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return `${day} ${BULAN_INDONESIA[month - 1]} ${year}`;
    }
  }

  // 5. Standard Indonesian / UK format: DD/MM/YYYY or DD-MM-YYYY (e.g. 01/01/2025 or 1-1-2025)
  const dmyMatch = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (dmyMatch) {
    const day = parseInt(dmyMatch[1], 10);
    const month = parseInt(dmyMatch[2], 10);
    const year = parseInt(dmyMatch[3], 10);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return `${day} ${BULAN_INDONESIA[month - 1]} ${year}`;
    }
  }

  // 6. Check if standard Date parser can parse it (e.g. "Jan 1, 2025" or "2025-01-01T00:00:00Z")
  const parsedDate = new Date(str);
  if (!isNaN(parsedDate.getTime()) && parsedDate.getFullYear() > 1900) {
    return `${parsedDate.getDate()} ${BULAN_INDONESIA[parsedDate.getMonth()]} ${parsedDate.getFullYear()}`;
  }

  return str;
};

// Helper: Format raw numeric/text value to Indonesian Rupiah (e.g. 15000000 -> Rp 15.000.000)
const formatRupiah = (value) => {
  if (value === null || value === undefined || value === '') return '';
  const str = value.toString().trim();
  if (!str) return '';

  if (/^rp/i.test(str)) return str;

  // Clean numbers
  const cleaned = str.replace(/[^0-9,-]/g, '');
  if (!cleaned) return str;

  const parts = cleaned.split(',');
  const integerPart = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  const decimalPart = parts.length > 1 ? `,${parts[1]}` : '';

  return `Rp ${integerPart}${decimalPart}`;
};

// Helper: Format live typing for nominal inputs
const formatNominalTyping = (value) => {
  if (!value) return '';
  const digits = value.replace(/[^0-9]/g, '');
  if (!digits) return '';
  return 'Rp ' + digits.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
};

function App() {
  const [templateFile, setTemplateFile] = useState(null);
  const [recipients, setRecipients] = useState([
    { id: 'rec-1', name: '', nominal: '' }
  ]);
  const [manualText, setManualText] = useState('');
  const [inputTab, setInputTab] = useState('excel'); // 'excel' | 'paste' | 'table'
  const [isProcessing, setIsProcessing] = useState(false);
  const [hasGenerated, setHasGenerated] = useState(false);
  const [downloadData, setDownloadData] = useState({ blob: null, fileName: '', isZip: false });
  const [activeStep, setActiveStep] = useState(1);
  const [showConfetti, setShowConfetti] = useState(false);
  const [recipientListRef] = useAutoAnimate();
  const fileInputRef = useRef(null);
  const { width, height } = useWindowSize();

  const [formData, setFormData] = useState({
    Kota: '',
    Tanggal_Konfirmasi: '',
    Periode: '',
    Nama_Klien: '',
    Sebutan1: '',
    Auditor1: '',
    Sebutan2: '',
    Auditor2: '',
    Tanggal_Jatuh_Tempo: '',
    Nama_Direktur: '',
    Jabatan: '',
    Nominal_Default: ''
  });

  const handleInputChange = (e) => {
    const { name, value } = e.target;
    setFormData((prev) => ({ ...prev, [name]: value }));
  };

  // Helper to add a new recipient row
  const addRecipientRow = () => {
    setRecipients((prev) => [
      ...prev,
      { id: `rec-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`, name: '', nominal: '' }
    ]);
  };

  // Helper to update a recipient field
  const updateRecipient = (id, field, value) => {
    setRecipients((prev) =>
      prev.map((r) => {
        if (r.id === id) {
          return {
            ...r,
            [field]: field === 'nominal' ? (value ? formatNominalTyping(value) : '') : value
          };
        }
        return r;
      })
    );
  };

  // Helper to remove a recipient row
  const removeRecipient = (id) => {
    setRecipients((prev) => {
      const filtered = prev.filter((r) => r.id !== id);
      return filtered.length > 0 ? filtered : [{ id: `rec-${Date.now()}`, name: '', nominal: '' }];
    });
  };

  // Helper to clear all recipients
  const clearAllRecipients = () => {
    if (window.confirm('Hapus semua daftar penerima?')) {
      setRecipients([{ id: `rec-${Date.now()}`, name: '', nominal: '' }]);
      setManualText('');
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  // Excel Upload Handler with smart column and date detection
  const handleExcelUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheet = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheet];
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

        if (!json || json.length === 0) {
          alert('File Excel kosong atau tidak memiliki data baris!');
          return;
        }

        const sampleRow = json[0];
        const keys = Object.keys(sampleRow);

        // Find best match for Name column
        const nameKey = keys.find((k) => {
          const lower = k.toLowerCase();
          return (
            lower.includes('nama') ||
            lower.includes('debitur') ||
            lower.includes('penerima') ||
            lower.includes('klien') ||
            lower.includes('customer') ||
            lower.includes('perusahaan')
          );
        }) || keys[0];

        // Find best match for Nominal column
        const nominalKey = keys.find((k) => {
          const lower = k.toLowerCase();
          return (
            lower.includes('nominal') ||
            lower.includes('saldo') ||
            lower.includes('jumlah') ||
            lower.includes('piutang') ||
            lower.includes('nilai') ||
            lower.includes('amount') ||
            lower.includes('total') ||
            lower.includes('rp') ||
            lower.includes('tagihan')
          );
        });

        // Find match for Date columns if present in Excel
        const tanggalKonfirmasiKey = keys.find((k) => {
          const lower = k.toLowerCase();
          return lower.includes('tanggal_konfirmasi') || lower.includes('tgl_konfirmasi') || lower.includes('tgl_surat');
        });

        const periodeKey = keys.find((k) => {
          const lower = k.toLowerCase();
          return lower.includes('periode') || lower.includes('per_tanggal') || lower.includes('cut_off') || lower.includes('cutoff');
        });

        const jatuhTempoKey = keys.find((k) => {
          const lower = k.toLowerCase();
          return lower.includes('jatuh_tempo') || lower.includes('due_date') || lower.includes('batas_waktu');
        });

        // If Excel contains global audit details in first row, auto-fill formData if empty
        setFormData((prev) => ({
          ...prev,
          Tanggal_Konfirmasi: tanggalKonfirmasiKey && sampleRow[tanggalKonfirmasiKey]
            ? formatIndonesianDate(sampleRow[tanggalKonfirmasiKey])
            : prev.Tanggal_Konfirmasi,
          Periode: periodeKey && sampleRow[periodeKey]
            ? formatIndonesianDate(sampleRow[periodeKey])
            : prev.Periode,
          Tanggal_Jatuh_Tempo: jatuhTempoKey && sampleRow[jatuhTempoKey]
            ? formatIndonesianDate(sampleRow[jatuhTempoKey])
            : prev.Tanggal_Jatuh_Tempo
        }));

        const parsed = [];
        json.forEach((row, idx) => {
          const nameVal = row[nameKey] ? row[nameKey].toString().trim() : '';
          let nominalVal = nominalKey && row[nominalKey] !== undefined && row[nominalKey] !== ''
            ? row[nominalKey].toString().trim()
            : '';

          if (nominalVal && !/^rp/i.test(nominalVal) && !isNaN(Number(nominalVal.replace(/[,.]/g, '')))) {
            nominalVal = formatRupiah(nominalVal);
          }

          if (nameVal) {
            parsed.push({
              id: `excel-${Date.now()}-${idx}-${Math.random().toString(36).substr(2, 4)}`,
              name: nameVal,
              nominal: nominalVal
            });
          }
        });

        if (parsed.length === 0) {
          alert('Tidak ditemukan data nama pada file Excel!');
          return;
        }

        setRecipients(parsed);
        setActiveStep(3);
        setInputTab('table');
        alert(`Berhasil memuat ${parsed.length} penerima ${nominalKey ? `beserta kolom nominal (${nominalKey})` : ''} dari Excel!`);
      } catch (err) {
        console.error('Excel Import Error:', err);
        alert('Gagal membaca file Excel: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Parse batch text from textarea (e.g. copied from Excel/Sheets or typed with delimiters)
  const handleApplyBatchText = () => {
    if (!manualText.trim()) {
      alert('Harap masukkan teks daftar penerima terlebih dahulu!');
      return;
    }

    const lines = manualText.split('\n').map((l) => l.trim()).filter(Boolean);
    const parsed = lines.map((line, idx) => {
      let name = line;
      let nominal = '';

      if (line.includes('\t')) {
        const parts = line.split('\t');
        name = parts[0].trim();
        nominal = parts.slice(1).join(' ').trim();
      } else if (line.includes('|')) {
        const parts = line.split('|');
        name = parts[0].trim();
        nominal = parts.slice(1).join(' ').trim();
      } else if (line.includes(';')) {
        const parts = line.split(';');
        name = parts[0].trim();
        nominal = parts.slice(1).join(' ').trim();
      } else if (line.includes(',') && /[0-9]|rp/i.test(line.split(',').slice(1).join(','))) {
        const commaIdx = line.lastIndexOf(',');
        name = line.substring(0, commaIdx).trim();
        nominal = line.substring(commaIdx + 1).trim();
      }

      if (nominal && !/^rp/i.test(nominal) && !isNaN(Number(nominal.replace(/[,.]/g, '')))) {
        nominal = formatRupiah(nominal);
      }

      return {
        id: `batch-${Date.now()}-${idx}-${Math.random().toString(36).substr(2, 4)}`,
        name,
        nominal
      };
    });

    if (parsed.length > 0) {
      setRecipients(parsed);
      setInputTab('table');
      setActiveStep(3);
      alert(`${parsed.length} penerima berhasil dimuat ke tabel!`);
    }
  };

  // Valid recipients for generation
  const validRecipients = recipients.filter((r) => r.name && r.name.trim().length > 0);

  const generateDocuments = async () => {
    if (!templateFile) {
      alert('Harap upload file Template Word terlebih dahulu!');
      return;
    }

    if (validRecipients.length === 0) {
      alert('Harap masukkan setidaknya satu Nama Penerima!');
      return;
    }

    setIsProcessing(true);
    try {
      const reader = new FileReader();
      reader.onload = async (event) => {
        try {
          const content = event.target.result;
          const zipResult = new JSZip();

          // Formatted Indonesian Dates
          const formattedTanggalKonfirmasi = formatIndonesianDate(formData.Tanggal_Konfirmasi);
          const formattedPeriode = formatIndonesianDate(formData.Periode);
          const formattedTanggalJatuhTempo = formatIndonesianDate(formData.Tanggal_Jatuh_Tempo);

          validRecipients.forEach((rec, index) => {
            const zipTemplate = new PizZip(content);
            const doc = new Docxtemplater(zipTemplate, {
              paragraphLoop: true,
              linebreaks: true,
              delimiters: { start: '{{', end: '}}' }
            });

            const recipientNominal = rec.nominal?.trim() || formData.Nominal_Default?.trim() || '-';

            // Comprehensive mapping for case sensitivity, dates, and template variations
            const docData = {
              ...formData,
              // Dates formatted as "1 Januari 2025"
              Tanggal_Konfirmasi: formattedTanggalKonfirmasi,
              tanggal_konfirmasi: formattedTanggalKonfirmasi,
              TANGGAL_KONFIRMASI: formattedTanggalKonfirmasi,
              Tanggal: formattedTanggalKonfirmasi,
              tanggal: formattedTanggalKonfirmasi,
              TANGGAL: formattedTanggalKonfirmasi,
              Periode: formattedPeriode,
              periode: formattedPeriode,
              PERIODE: formattedPeriode,
              Tanggal_Jatuh_Tempo: formattedTanggalJatuhTempo,
              tanggal_jatuh_tempo: formattedTanggalJatuhTempo,
              TANGGAL_JATUH_TEMPO: formattedTanggalJatuhTempo,
              Jatuh_Tempo: formattedTanggalJatuhTempo,
              jatuh_tempo: formattedTanggalJatuhTempo,
              JATUH_TEMPO: formattedTanggalJatuhTempo,

              // Recipient Name
              Nama_Penerima: rec.name.trim(),
              nama_penerima: rec.name.trim(),
              NAMA_PENERIMA: rec.name.trim(),

              // Nominal Piutang
              nominal: recipientNominal,
              Nominal: recipientNominal,
              NOMINAL: recipientNominal,
              nominal_piutang: recipientNominal,
              Nominal_Piutang: recipientNominal,
              NOMINAL_PIUTANG: recipientNominal,
              saldo: recipientNominal,
              Saldo: recipientNominal,
              SALDO: recipientNominal,
              jumlah: recipientNominal,
              Jumlah: recipientNominal,
              JUMLAH: recipientNominal
            };

            // Keep empty fields as placeholder tag, or blank for titles
            Object.keys(docData).forEach((key) => {
              if (docData[key] === '' || docData[key] === undefined || docData[key] === null) {
                docData[key] = `{{${key}}}`;
              }
              if ((key === 'Sebutan1' || key === 'Sebutan2') && docData[key] === `{{${key}}}`) {
                docData[key] = '';
              }
            });

            doc.render(docData);
            const out = doc.getZip().generate({
              type: 'blob',
              mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            });

            const safeRecipientName = rec.name.replace(/[/\\?%*:|"<>]/g, '-').trim() || `Penerima-${index + 1}`;

            if (validRecipients.length === 1) {
              setDownloadData({
                blob: out,
                fileName: `Konfirmasi Piutang - ${safeRecipientName}.docx`,
                isZip: false
              });
            } else {
              zipResult.file(`Konfirmasi Piutang - ${safeRecipientName}.docx`, out);
            }
          });

          if (validRecipients.length > 1) {
            const zipContent = await zipResult.generateAsync({ type: 'blob' });
            setDownloadData({
              blob: zipContent,
              fileName: `Konfirmasi Piutang - ${formData.Nama_Klien || 'Klien'}.zip`,
              isZip: true
            });
          }

          setHasGenerated(true);
          setShowConfetti(true);
          setTimeout(() => setShowConfetti(false), 5000);
        } catch (error) {
          console.error('Error Detail:', error);
          if (error.properties && error.properties.errors instanceof Array) {
            const errorMessages = error.properties.errors.map((err) => `- ${err.properties.explanation}`).join('\n');
            alert('Sistem menemukan masalah pada template Word Anda:\n' + errorMessages);
          } else {
            alert('Terjadi kesalahan saat memproses dokumen: ' + error.message);
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

  const currentStep = hasGenerated ? 4 : activeStep;

  return (
    <div className="app-wrapper">
      {/* Header */}
      <motion.header
        className="app-header"
        initial={{ opacity: 0, y: -20 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.6, ease: 'easeOut' }}
      >
        <img src={logoTransparan} alt="Logo KAP" className="app-header__logo-img" />
        <div className="app-header__kap-badge">
          <Icons.Document /> Kantor Akuntan Publik
        </div>
        <h1 className="app-header__title">Generator Konfirmasi Piutang</h1>
        <p className="app-header__subtitle">
          Buat surat konfirmasi piutang untuk audit dengan cepat, mudah, dan otomatis
        </p>
      </motion.header>

      {/* Main Card */}
      <motion.main
        className="app-card"
        initial={{ opacity: 0, y: 30 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.5, delay: 0.2, ease: 'easeOut' }}
      >
        {/* Progress Bar */}
        <div className="progress-bar">
          {[
            { step: 1, label: 'Template' },
            { step: 2, label: 'Detail Audit' },
            { step: 3, label: 'Penerima & Nominal' },
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

        <AnimatePresence mode="wait">
          {!hasGenerated ? (
            <motion.div
              key="form"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0, y: -20 }}
              transition={{ duration: 0.3 }}
            >
              {/* Step 1: Template */}
              <motion.section
                className="section"
                initial={{ opacity: 0, x: -20 }}
                animate={{ opacity: 1, x: 0 }}
                transition={{ duration: 0.4, delay: 0.1 }}
              >
                <div className="section__header">
                  <div className="section__number">1</div>
                  <div>
                    <h3 className="section__title">Upload Template Word</h3>
                    <p className="section__description">Pilih file template .docx yang memuat placeholder konfirmasi piutang</p>
                  </div>
                </div>

                <div className="file-upload mt-3">
                  <motion.label
                    className="file-upload__area"
                    whileHover={{ scale: 1.01 }}
                    whileTap={{ scale: 0.99 }}
                  >
                    <input
                      type="file"
                      accept=".docx"
                      className="file-upload__input"
                      onChange={(e) => {
                        if (e.target.files && e.target.files[0]) {
                          setTemplateFile(e.target.files[0]);
                          setActiveStep(2);
                        }
                      }}
                    />
                    <div className="file-upload__icon"><Icons.Upload /></div>
                    <p className="file-upload__text">
                      <strong>Pilih file template</strong> atau drag & drop di sini
                    </p>
                    <p className="form-hint mt-2">Mendukung tag: <code>&#123;&#123;Nama_Penerima&#125;&#125;</code>, <code>&#123;&#123;nominal&#125;&#125;</code>, <code>&#123;&#123;Periode&#125;&#125;</code>, dll.</p>
                  </motion.label>

                  {templateFile && (
                    <div className="file-upload__preview">
                      <span className="file-upload__preview-icon"><Icons.File /></span>
                      <span style={{ fontWeight: 600 }}>{templateFile.name}</span>
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
              </motion.section>

              {/* Step 2: Detail Audit */}
              <motion.section
                className="section"
                initial={{ opacity: 0, x: -20 }}
                animate={{ opacity: 1, x: 0 }}
                transition={{ duration: 0.4, delay: 0.2 }}
              >
                <div className="section__header">
                  <div className="section__number">2</div>
                  <div>
                    <h3 className="section__title">Detail Audit</h3>
                    <p className="section__description">Lengkapi informasi umum audit yang akan tercantum di surat</p>
                  </div>
                </div>

                <div className="form-grid mt-3">
                  <div className="form-group">
                    <label className="form-label">Kota Surat <span className="form-label__required">*</span></label>
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
                    {formData.Tanggal_Konfirmasi && (
                      <p className="form-hint" style={{ color: '#059669', fontWeight: 600 }}>
                        <Icons.Calendar /> Format di surat: {formatIndonesianDate(formData.Tanggal_Konfirmasi)}
                      </p>
                    )}
                  </div>

                  <div className="form-group">
                    <label className="form-label">Periode Audit <span className="form-label__required">*</span></label>
                    <input
                      name="Periode"
                      className="form-input"
                      placeholder="31 Desember 2024 atau 2024-12-31"
                      onChange={handleInputChange}
                      value={formData.Periode}
                    />
                    {formData.Periode && (
                      <p className="form-hint" style={{ color: '#059669', fontWeight: 600 }}>
                        <Icons.Calendar /> Format di surat: {formatIndonesianDate(formData.Periode)}
                      </p>
                    )}
                  </div>

                  <div className="form-group">
                    <label className="form-label">Nama Klien <span className="form-label__required">*</span></label>
                    <input
                      name="Nama_Klien"
                      className="form-input"
                      placeholder="PT Contoh Sejahtera Abadi"
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
                        placeholder="Nama Auditor 1"
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
                        placeholder="Nama Auditor 2"
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
                    {formData.Tanggal_Jatuh_Tempo && (
                      <p className="form-hint" style={{ color: '#059669', fontWeight: 600 }}>
                        <Icons.Calendar /> Format di surat: {formatIndonesianDate(formData.Tanggal_Jatuh_Tempo)}
                      </p>
                    )}
                  </div>

                  <div className="form-row">
                    <div className="form-row__item">
                      <label className="form-label">Nama Direktur / Penandatangan <span className="form-label__required">*</span></label>
                      <input
                        name="Nama_Direktur"
                        className="form-input"
                        placeholder="Nama Direktur Klien"
                        onChange={handleInputChange}
                        value={formData.Nama_Direktur}
                      />
                    </div>
                    <div className="form-row__item" style={{ maxWidth: '180px' }}>
                      <label className="form-label">Jabatan</label>
                      <input
                        name="Jabatan"
                        className="form-input"
                        placeholder="Direktur Utama"
                        onChange={handleInputChange}
                        value={formData.Jabatan}
                      />
                    </div>
                  </div>

                  <div className="form-group">
                    <label className="form-label">Nominal Piutang Standar / Fallback (Opsional)</label>
                    <input
                      name="Nominal_Default"
                      className="form-input"
                      placeholder="Contoh: Rp 10.000.000 (digunakan jika baris kosong)"
                      onChange={(e) => {
                        const val = e.target.value;
                        setFormData((prev) => ({
                          ...prev,
                          Nominal_Default: val ? formatNominalTyping(val) : ''
                        }));
                      }}
                      value={formData.Nominal_Default}
                    />
                    <p className="form-hint">Digunakan jika ada baris penerima yang tidak memiliki nilai nominal tersendiri.</p>
                  </div>
                </div>
              </motion.section>

              {/* Step 3: Daftar Penerima & Nominal */}
              <motion.section
                className="section"
                initial={{ opacity: 0, x: -20 }}
                animate={{ opacity: 1, x: 0 }}
                transition={{ duration: 0.4, delay: 0.3 }}
              >
                <div className="section__header">
                  <div className="section__number">3</div>
                  <div style={{ flex: 1 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
                      <h3 className="section__title">Daftar Penerima & Nominal Piutang</h3>
                      <span className="badge badge--primary">
                        <Icons.Users /> {validRecipients.length} Penerima
                      </span>
                    </div>
                    <p className="section__description">
                      Masukkan nama debitur dan nominal saldo piutang untuk dimasukkan ke <code>&#123;&#123;nominal&#125;&#125;</code>
                    </p>
                  </div>
                </div>

                {/* Input Method Switcher Tabs */}
                <div className="tab-switcher mt-3">
                  <button
                    type="button"
                    className={`tab-btn ${inputTab === 'excel' ? 'active' : ''}`}
                    onClick={() => setInputTab('excel')}
                  >
                    <Icons.Upload /> Import Excel
                  </button>
                  <button
                    type="button"
                    className={`tab-btn ${inputTab === 'paste' ? 'active' : ''}`}
                    onClick={() => setInputTab('paste')}
                  >
                    <Icons.Edit /> Paste Batch (Teks/Excel)
                  </button>
                  <button
                    type="button"
                    className={`tab-btn ${inputTab === 'table' ? 'active' : ''}`}
                    onClick={() => setInputTab('table')}
                  >
                    <Icons.Table /> Tabel Interaktif ({validRecipients.length})
                  </button>
                </div>

                {/* Tab 1: Import Excel */}
                {inputTab === 'excel' && (
                  <div className="mt-3">
                    <div className="file-upload">
                      <motion.label
                        className="file-upload__area file-upload__area--compact"
                        whileHover={{ scale: 1.01 }}
                        whileTap={{ scale: 0.99 }}
                      >
                        <input
                          ref={fileInputRef}
                          type="file"
                          accept=".xlsx, .xls"
                          className="file-upload__input"
                          onChange={handleExcelUpload}
                        />
                        <div className="file-upload__icon"><Icons.Users /></div>
                        <p className="file-upload__text">
                          <strong>Upload file Excel (.xlsx / .xls)</strong>
                        </p>
                        <p className="form-hint mt-1">
                          Mendukung kolom <strong>Nama</strong>, <strong>Nominal Piutang</strong>, dan kolom <strong>Tanggal</strong>
                        </p>
                      </motion.label>
                    </div>

                    <div className="guide-box mt-2">
                      <Icons.Info />
                      <div>
                        <strong>Tips Format Excel:</strong>
                        <p>Format tanggal pada Excel (seperti <code>2025-01-01</code> atau <code>01/01/2025</code>) akan otomatis diekstrak menjadi <code>1 Januari 2025</code>.</p>
                      </div>
                    </div>
                  </div>
                )}

                {/* Tab 2: Batch Paste Text */}
                {inputTab === 'paste' && (
                  <div className="mt-3">
                    <label className="form-label mb-2">Paste Data Penerima & Nominal</label>
                    <textarea
                      className="form-textarea"
                      style={{ minHeight: '120px', fontFamily: 'monospace', fontSize: '0.875rem' }}
                      placeholder={"PT Maju Bersama | 15000000\nCV Sinar Terang ; 25.000.000\nPT Berkah Abadi\t50000000\nToko Mulia Jaya, Rp 10.000.000"}
                      value={manualText}
                      onChange={(e) => setManualText(e.target.value)}
                    />
                    <p className="form-hint mt-1">
                      Format per baris: <code>Nama | Nominal</code> atau copy langsung 2 kolom dari tabel Excel / Google Sheets (otomatis dipisahkan Tab).
                    </p>
                    <button
                      type="button"
                      className="btn btn--outline mt-2"
                      onClick={handleApplyBatchText}
                    >
                      <Icons.Table /> Terapkan ke Tabel Penerima
                    </button>
                  </div>
                )}

                {/* Tab 3 & Global: Live Recipient Table */}
                <div className="recipient-table-container mt-4">
                  <div className="recipient-table-header">
                    <div className="recipient-table-title">
                      <Icons.Table /> <strong>Tabel Penerima & Nominal Piutang</strong>
                    </div>
                    <div className="recipient-table-actions">
                      <button
                        type="button"
                        className="btn btn--outline btn--sm"
                        onClick={addRecipientRow}
                      >
                        <Icons.Plus /> Tambah Baris
                      </button>
                      {recipients.length > 1 && (
                        <button
                          type="button"
                          className="btn btn--ghost btn--sm"
                          onClick={clearAllRecipients}
                        >
                          <Icons.Trash /> Bersihkan
                        </button>
                      )}
                    </div>
                  </div>

                  <div className="recipient-table-wrapper">
                    <table className="recipient-table">
                      <thead>
                        <tr>
                          <th style={{ width: '45px', textAlign: 'center' }}>No</th>
                          <th>Nama Penerima / Debitur <span className="form-label__required">*</span></th>
                          <th style={{ width: '220px' }}>Nominal Piutang (Rp)</th>
                          <th style={{ width: '50px', textAlign: 'center' }}>Aksi</th>
                        </tr>
                      </thead>
                      <tbody ref={recipientListRef}>
                        {recipients.map((rec, index) => (
                          <tr key={rec.id} className="recipient-row">
                            <td style={{ textAlign: 'center', fontWeight: 600, color: '#94a3b8' }}>
                              {index + 1}
                            </td>
                            <td>
                              <input
                                type="text"
                                className="form-input form-input--dense"
                                placeholder="Contoh: PT Sumber Rejeki"
                                value={rec.name}
                                onChange={(e) => updateRecipient(rec.id, 'name', e.target.value)}
                              />
                            </td>
                            <td>
                              <input
                                type="text"
                                className="form-input form-input--dense form-input--money"
                                placeholder="Rp 0"
                                value={rec.nominal}
                                onChange={(e) => updateRecipient(rec.id, 'nominal', e.target.value)}
                              />
                            </td>
                            <td style={{ textAlign: 'center' }}>
                              <button
                                type="button"
                                className="btn-icon btn-icon--danger"
                                title="Hapus baris"
                                onClick={() => removeRecipient(rec.id)}
                              >
                                <Icons.Trash />
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  <div className="recipient-table-footer">
                    <button
                      type="button"
                      className="btn btn--outline btn--sm"
                      onClick={addRecipientRow}
                    >
                      <Icons.Plus /> Tambah Penerima Baru
                    </button>
                    <span className="form-hint">
                      Total <strong>{validRecipients.length}</strong> penerima siap digenerate
                    </span>
                  </div>
                </div>
              </motion.section>

              {/* Action Bar */}
              <div className="action-bar">
                <button
                  type="button"
                  className="btn btn--ghost"
                  onClick={() => {
                    if (window.confirm('Reset semua formulir dan data penerima?')) {
                      setTemplateFile(null);
                      setRecipients([{ id: `rec-${Date.now()}`, name: '', nominal: '' }]);
                      setManualText('');
                      setFormData({
                        Kota: '',
                        Tanggal_Konfirmasi: '',
                        Periode: '',
                        Nama_Klien: '',
                        Sebutan1: '',
                        Auditor1: '',
                        Sebutan2: '',
                        Auditor2: '',
                        Tanggal_Jatuh_Tempo: '',
                        Nama_Direktur: '',
                        Jabatan: '',
                        Nominal_Default: ''
                      });
                      setActiveStep(1);
                    }
                  }}
                >
                  Reset Form
                </button>
                <motion.button
                  className="btn btn--primary btn--lg"
                  onClick={generateDocuments}
                  disabled={isProcessing || !templateFile || validRecipients.length === 0}
                  whileHover={{ scale: 1.03 }}
                  whileTap={{ scale: 0.97 }}
                >
                  {isProcessing ? (
                    <>
                      <span className="btn__spinner"></span>
                      Sedang Membuat Dokumen...
                    </>
                  ) : (
                    <>
                      <Icons.Sparkles />
                      Generate {validRecipients.length || 1} Surat Konfirmasi
                    </>
                  )}
                </motion.button>
              </div>
            </motion.div>
          ) : (
            <motion.div
              key="result"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              transition={{ duration: 0.3 }}
            >
              {showConfetti && (
                <ReactConfetti
                  width={width}
                  height={height}
                  recycle={false}
                  numberOfPieces={300}
                  colors={['#10b981', '#0d9488', '#059669', '#34d399', '#6ee7b7']}
                />
              )}

              {/* Result Card */}
              <motion.section
                className="section"
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                transition={{ duration: 0.5, ease: [0.34, 1.56, 0.64, 1] }}
              >
                <div className="result-card">
                  <motion.div
                    className="result-card__icon"
                    initial={{ scale: 0 }}
                    animate={{ scale: 1 }}
                    transition={{ delay: 0.2, duration: 0.5, ease: [0.34, 1.56, 0.64, 1] }}
                  >
                    <Icons.Check />
                  </motion.div>
                  <span className="result-card__badge"><Icons.Sparkles /> Selesai</span>
                  <h4 className="result-card__title">Dokumen Berhasil Dibuat!</h4>
                  <p className="result-card__message">
                    {downloadData.isZip
                      ? `${validRecipients.length} file surat konfirmasi piutang telah dikemas dalam arsip ZIP.`
                      : 'Surat konfirmasi piutang siap diunduh dalam format Word (.docx).'
                    }
                  </p>

                  <div className="result-card__actions">
                    <motion.button
                      className="btn btn--success btn--lg"
                      onClick={() => saveAs(downloadData.blob, downloadData.fileName)}
                      whileHover={{ scale: 1.03 }}
                      whileTap={{ scale: 0.97 }}
                    >
                      <Icons.Download />
                      Unduh {downloadData.isZip ? 'Semua Dokumen (ZIP)' : 'Dokumen Word'}
                    </motion.button>

                    <button
                      type="button"
                      className="btn btn--outline"
                      onClick={() => {
                        alert("💡 Panduan Penggunaan Dokumen:\n\n• Buka file dokumen di Microsoft Word.\n• Periksa kembali nominal piutang dan identitas penerima.\n• Anda dapat langsung mencetak atau 'Save As' ke PDF untuk dikirimkan ke pihak terkait.");
                      }}
                    >
                      <Icons.Info /> Panduan
                    </button>
                  </div>
                </div>

                <div className="action-bar">
                  <button
                    type="button"
                    className="btn btn--ghost"
                    onClick={() => {
                      setHasGenerated(false);
                      setActiveStep(3);
                    }}
                  >
                    <Icons.ArrowLeft /> Kembali Edit Data
                  </button>
                  <button
                    type="button"
                    className="btn btn--outline"
                    onClick={() => {
                      setHasGenerated(false);
                      setActiveStep(1);
                      setTemplateFile(null);
                      setRecipients([{ id: `rec-${Date.now()}`, name: '', nominal: '' }]);
                      setManualText('');
                    }}
                  >
                    <Icons.Sparkles /> Buat Konfirmasi Baru
                  </button>
                </div>
              </motion.section>
            </motion.div>
          )}
        </AnimatePresence>
      </motion.main>

      {/* Footer */}
      <footer className="text-center text-muted" style={{ fontSize: '0.8125rem' }}>
        <p>Generator Konfirmasi Piutang • Kantor Akuntan Publik</p>
      </footer>
    </div>
  );
}

export default App;

