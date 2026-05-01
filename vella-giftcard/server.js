require('dotenv').config();
const express    = require('express');
const path       = require('path');
const fs         = require('fs');
const jwt        = require('jsonwebtoken');
const bcrypt     = require('bcryptjs');
const axios      = require('axios');
const QRCode     = require('qrcode');
const { v4: uuidv4 } = require('uuid');
const multer     = require('multer');
const nodemailer = require('nodemailer');

const app      = express();
const PORT     = process.env.PORT || 3002;
const BASE_URL = process.env.BASE_URL || `http://localhost:${PORT}`;
const SECRET   = process.env.JWT_SECRET || 'vella-gift-secret';

app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

// ── JSON database ────────────────────────────────────────────────────────────
const DB_DIR = path.join(__dirname, 'db');
['db', 'uploads'].forEach(d => {
  const p = path.join(__dirname, d);
  if (!fs.existsSync(p)) fs.mkdirSync(p, { recursive: true });
});

const db = {
  read:   k      => { try { return JSON.parse(fs.readFileSync(path.join(DB_DIR, k+'.json'), 'utf8')); } catch { return []; } },
  write:  (k, v) => fs.writeFileSync(path.join(DB_DIR, k+'.json'), JSON.stringify(v, null, 2)),
  find:   (k, f) => db.read(k).find(f),
  filter: (k, f) => f ? db.read(k).filter(f) : db.read(k),
  insert: (k, v) => { const a = db.read(k); a.push(v); db.write(k, a); return v; },
  update: (k, f, ch) => {
    const a = db.read(k), i = a.findIndex(f);
    if (i < 0) return null;
    a[i] = { ...a[i], ...ch }; db.write(k, a); return a[i];
  },
  remove: (k, f) => db.write(k, db.read(k).filter(x => !f(x))),
};

// ── Seed ─────────────────────────────────────────────────────────────────────
(async () => {
  if (!db.find('users', u => u.email === 'diogo.hhonda@gmail.com')) {
    db.insert('users', {
      id: uuidv4(), name: 'Diogo Honda',
      email: 'diogo.hhonda@gmail.com',
      password: await bcrypt.hash('Diogo9968', 10),
      role: 'admin', createdAt: new Date().toISOString()
    });
    console.log('✅ Admin criado: diogo.hhonda@gmail.com / Diogo9968');
  }
  if (db.read('templates').length === 0) {
    [
      { name: 'Aniversário Especial',    value: 50,  desc: 'Uma surpresa para este dia especial',     color: 'linear-gradient(135deg,#7C3AED,#EC4899)' },
      { name: 'Presente de Natal',       value: 100, desc: 'Com muito carinho neste Natal',           color: 'linear-gradient(135deg,#DC2626,#166534)' },
      { name: 'Parabéns pela conquista', value: 150, desc: 'Você merece! Comemore muito!',            color: 'linear-gradient(135deg,#D97706,#92400E)' },
    ].forEach(t => db.insert('templates', { id: uuidv4(), ...t, img: null, active: true, createdAt: new Date().toISOString() }));
  }
})();

// ── Auth middleware ───────────────────────────────────────────────────────────
const auth = (req, res, next) => {
  const t = req.headers.authorization?.split(' ')[1];
  if (!t) return res.status(401).json({ error: 'Token necessário' });
  try { req.user = jwt.verify(t, SECRET); next(); }
  catch { res.status(401).json({ error: 'Token inválido' }); }
};
const adminOnly = (req, res, next) =>
  req.user?.role === 'admin' ? next() : res.status(403).json({ error: 'Acesso negado' });

// ── Upload ────────────────────────────────────────────────────────────────────
const upload = multer({
  storage: multer.diskStorage({
    destination: path.join(__dirname, 'uploads'),
    filename: (_, f, cb) => cb(null, Date.now() + path.extname(f.originalname))
  }),
  limits: { fileSize: 5 * 1024 * 1024 },
  fileFilter: (_, f, cb) => f.mimetype.startsWith('image/') ? cb(null, true) : cb(new Error('Apenas imagens'))
});

// ── Asaas helper ──────────────────────────────────────────────────────────────
const ASAAS_BASE = process.env.ASAAS_ENV === 'prod'
  ? 'https://api.asaas.com/v3'
  : 'https://sandbox.asaas.com/api/v3';

const asaas = axios.create({
  baseURL: ASAAS_BASE,
  headers: { 'access_token': process.env.ASAAS_KEY || '', 'Content-Type': 'application/json' }
});

async function createPixPayment(buyer, value, desc, ref) {
  try {
    let customerId;
    try {
      const r = await asaas.post('/customers', { name: buyer.name, email: buyer.email, notificationDisabled: true });
      customerId = r.data.id;
    } catch {
      const r = await asaas.get(`/customers?email=${encodeURIComponent(buyer.email)}&limit=1`);
      customerId = r.data.data?.[0]?.id;
      if (!customerId) throw new Error('Cliente não encontrado no Asaas');
    }
    const due = new Date(Date.now() + 86400000).toISOString().split('T')[0];
    const pay = await asaas.post('/payments', { customer: customerId, billingType: 'PIX', value, dueDate: due, description: desc, externalReference: ref });
    const qr  = await asaas.get(`/payments/${pay.data.id}/pixQrCode`);
    return { paymentId: pay.data.id, pixKey: qr.data.payload, pixQRImage: qr.data.encodedImage };
  } catch (e) {
    console.warn('Asaas indisponível, usando modo demo:', e.response?.data?.errors?.[0]?.description || e.message);
    return null;
  }
}

// QR code para Pix manual (fallback sem Asaas)
async function buildManualPixQR(value, ref) {
  const pixKey  = process.env.PIX_KEY || 'diogo.hhonda@gmail.com';
  const name    = 'Vella Gift';
  const city    = 'SAO PAULO';
  const txid    = ref.replace(/[^A-Za-z0-9]/g, '').slice(0, 25);
  const amount  = value.toFixed(2);

  function crc16(str) {
    let crc = 0xFFFF;
    for (let i = 0; i < str.length; i++) {
      crc ^= str.charCodeAt(i) << 8;
      for (let j = 0; j < 8; j++) crc = (crc & 0x8000) ? (crc << 1) ^ 0x1021 : crc << 1;
    }
    return (crc & 0xFFFF).toString(16).toUpperCase().padStart(4, '0');
  }
  function field(id, v) { return id + String(v.length).padStart(2,'0') + v; }

  const merchantInfo = field('00', 'BR.GOV.BCB.PIX') + field('01', pixKey);
  const payload = [
    field('00', '01'),
    field('26', merchantInfo),
    field('52', '0000'),
    field('53', '986'),
    field('54', amount),
    field('58', 'BR'),
    field('59', name.slice(0,25)),
    field('60', city),
    field('62', field('05', txid)),
    '6304'
  ].join('');
  const final = payload + crc16(payload);
  return { pixKey: final, pixQRImage: await QRCode.toDataURL(final, { width: 280, errorCorrectionLevel: 'M', color: { dark: '#1E1B4B', light: '#ffffff' } }) };
}

// ── E-mail helper ─────────────────────────────────────────────────────────────
async function sendEmail(to, benefName, cardLink, value, message, qrDataUrl) {
  if (!process.env.SMTP_HOST) { console.log(`📧 [sem SMTP] Link do cartão: ${cardLink}`); return; }
  const t = nodemailer.createTransport({ host: process.env.SMTP_HOST, port: +process.env.SMTP_PORT || 587, secure: false, auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS } });
  await t.sendMail({
    from: `"Vella Gift" <${process.env.SMTP_USER}>`,
    to,
    subject: `🎁 Você recebeu um Gift Card Vella — R$ ${value}`,
    html: `<div style="font-family:sans-serif;max-width:520px;margin:auto;padding:24px">
      <h2 style="color:#7C3AED">🎁 Vella Gift Card — R$ ${value}</h2>
      <p>Olá <strong>${benefName}</strong>, você recebeu um presente!</p>
      ${message ? `<blockquote style="border-left:4px solid #7C3AED;padding-left:12px;color:#555;font-style:italic">"${message}"</blockquote>` : ''}
      <p>Clique para ver e adicionar à sua carteira digital:</p>
      <a href="${cardLink}" style="display:inline-block;background:linear-gradient(135deg,#7C3AED,#EC4899);color:#fff;padding:14px 28px;border-radius:10px;text-decoration:none;font-weight:700;margin:12px 0">
        Ver meu Gift Card →
      </a>
      <p style="margin-top:16px;color:#666;font-size:13px">Ou escaneie o QR Code:</p>
      <img src="${qrDataUrl}" style="width:180px;display:block;margin:8px 0">
      <hr style="margin:24px 0;border:none;border-top:1px solid #eee">
      <p style="font-size:11px;color:#aaa">Vella Kids Gift Cards — Presentes que fazem a diferença</p>
    </div>`
  });
}

// ── ROTAS ─────────────────────────────────────────────────────────────────────

// Auth
app.post('/api/auth/login', async (req, res) => {
  const { email, password } = req.body;
  const u = db.find('users', u => u.email.toLowerCase() === email?.trim().toLowerCase());
  if (!u || !(await bcrypt.compare(password, u.password)))
    return res.status(401).json({ error: 'E-mail ou senha inválidos' });
  const token = jwt.sign({ id: u.id, name: u.name, email: u.email, role: u.role }, SECRET, { expiresIn: '30d' });
  res.json({ token, user: { id: u.id, name: u.name, email: u.email, role: u.role } });
});

app.post('/api/auth/register', async (req, res) => {
  const { name, email, password } = req.body;
  if (!name || !email || !password) return res.status(400).json({ error: 'Todos os campos são obrigatórios' });
  if (password.length < 6) return res.status(400).json({ error: 'Senha mínima: 6 caracteres' });
  if (db.find('users', u => u.email.toLowerCase() === email.toLowerCase()))
    return res.status(400).json({ error: 'E-mail já cadastrado' });
  const u = db.insert('users', { id: uuidv4(), name: name.trim(), email: email.trim().toLowerCase(), password: await bcrypt.hash(password, 10), role: 'client', createdAt: new Date().toISOString() });
  const token = jwt.sign({ id: u.id, name: u.name, email: u.email, role: 'client' }, SECRET, { expiresIn: '30d' });
  res.json({ token, user: { id: u.id, name: u.name, email: u.email, role: 'client' } });
});

app.put('/api/auth/profile', auth, async (req, res) => {
  const { name, password } = req.body;
  const ch = {};
  if (name?.trim()) ch.name = name.trim();
  if (password?.length >= 6) ch.password = await bcrypt.hash(password, 10);
  else if (password?.length > 0) return res.status(400).json({ error: 'Senha mínima: 6 caracteres' });
  db.update('users', u => u.id === req.user.id, ch);
  res.json({ ok: true });
});

// Templates
app.get('/api/templates', (_, res) => res.json(db.filter('templates', t => t.active)));

app.post('/api/templates', auth, adminOnly, upload.single('img'), (req, res) => {
  const { name, value, desc, color } = req.body;
  if (!name || !value || !desc) return res.status(400).json({ error: 'Campos obrigatórios: nome, valor, descrição' });
  const img = req.file ? `/uploads/${req.file.filename}` : null;
  res.json(db.insert('templates', { id: uuidv4(), name: name.trim(), value: parseFloat(value), desc: desc.trim(), color: color || 'linear-gradient(135deg,#7C3AED,#EC4899)', img, active: true, createdAt: new Date().toISOString() }));
});

app.delete('/api/templates/:id', auth, adminOnly, (req, res) => {
  db.remove('templates', t => t.id === req.params.id);
  res.json({ ok: true });
});

// Pedidos
app.post('/api/orders', auth, async (req, res) => {
  const { templateId, benefName, benefEmail, message } = req.body;
  if (!templateId || !benefName?.trim()) return res.status(400).json({ error: 'Template e nome do beneficiário são obrigatórios' });

  const tmpl = db.find('templates', t => t.id === templateId && t.active);
  if (!tmpl) return res.status(404).json({ error: 'Template não encontrado' });

  const code     = 'VG' + Math.random().toString(36).slice(2, 9).toUpperCase();
  const cardLink = `${BASE_URL}/card/${code}`;
  const cardQR   = await QRCode.toDataURL(cardLink, { width: 256, errorCorrectionLevel: 'H', color: { dark: '#1E1B4B', light: '#ffffff' } });

  // Tenta Asaas; cai em demo se não configurado
  let payment = null;
  if (process.env.ASAAS_KEY) {
    payment = await createPixPayment(req.user, tmpl.value, `Gift Card Vella — ${tmpl.name}`, code);
  }
  if (!payment) {
    payment = await buildManualPixQR(tmpl.value, code);
    payment.paymentId = null;
  }

  const order = db.insert('orders', {
    id: uuidv4(), code, templateId,
    buyerId: req.user.id, buyerName: req.user.name, buyerEmail: req.user.email,
    benefName: benefName.trim(), benefEmail: benefEmail?.trim() || null,
    message: message?.trim() || '',
    value: tmpl.value, templateName: tmpl.name, color: tmpl.color, img: tmpl.img,
    paymentId: payment.paymentId, status: 'pending',
    cardLink, createdAt: new Date().toISOString()
  });

  res.json({ ...order, cardQR, pixKey: payment.pixKey, pixQRImage: payment.pixQRImage });
});

// Admin confirma pagamento manualmente (modo demo / sandbox)
app.post('/api/orders/:code/confirm', auth, adminOnly, async (req, res) => {
  const order = db.update('orders', o => o.code === req.params.code, { status: 'active', paidAt: new Date().toISOString() });
  if (!order) return res.status(404).json({ error: 'Pedido não encontrado' });
  // Envia e-mail ao beneficiário se tiver e-mail e SMTP configurados
  if (order.benefEmail) {
    const qr = await QRCode.toDataURL(order.cardLink, { width: 200 });
    sendEmail(order.benefEmail, order.benefName, order.cardLink, order.value, order.message, qr).catch(console.warn);
  }
  res.json(order);
});

app.get('/api/orders',      auth, adminOnly, (_, res) => res.json(db.filter('orders')));
app.get('/api/orders/mine', auth, (req, res) => res.json(db.filter('orders', o => o.buyerId === req.user.id)));

app.patch('/api/orders/:code/use', auth, adminOnly, (req, res) => {
  const o = db.update('orders', o => o.code === req.params.code, { status: 'used', usedAt: new Date().toISOString() });
  o ? res.json(o) : res.status(404).json({ error: 'Pedido não encontrado' });
});

// Cartão público (beneficiário)
app.get('/api/cards/:code', async (req, res) => {
  const order = db.find('orders', o => o.code === req.params.code);
  if (!order) return res.status(404).json({ error: 'Gift Card não encontrado' });
  const qr = await QRCode.toDataURL(order.cardLink, { width: 200, errorCorrectionLevel: 'H', color: { dark: '#1E1B4B', light: '#ffffff' } });
  res.json({ ...order, qr });
});

// Admin: clientes e stats
app.get('/api/clients', auth, adminOnly, (_, res) =>
  res.json(db.filter('users', u => u.role !== 'admin').map(({ password, ...u }) => u)));

app.get('/api/stats', auth, adminOnly, (_, res) => {
  const orders    = db.filter('orders');
  const revenue   = orders.filter(o => ['active','used'].includes(o.status)).reduce((s, o) => s + o.value, 0);
  const pending   = orders.filter(o => o.status === 'pending').length;
  res.json({ orders: orders.length, revenue, pending, clients: db.filter('users', u => u.role !== 'admin').length, templates: db.filter('templates', t => t.active).length });
});

// Webhook Asaas
app.post('/api/webhook/asaas', express.text({ type: '*/*' }), async (req, res) => {
  try {
    const ev = JSON.parse(req.body);
    if (['PAYMENT_RECEIVED','PAYMENT_CONFIRMED'].includes(ev.event)) {
      const order = db.find('orders', o => o.paymentId === ev.payment?.id);
      if (order) {
        db.update('orders', o => o.id === order.id, { status: 'active', paidAt: new Date().toISOString() });
        if (order.benefEmail) {
          const qr = await QRCode.toDataURL(order.cardLink, { width: 200 });
          sendEmail(order.benefEmail, order.benefName, order.cardLink, order.value, order.message, qr).catch(console.warn);
        }
      }
    }
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// QR code como imagem PNG (para usar em <img src="/api/qr/CODE">)
app.get('/api/qr/:code', async (req, res) => {
  const order = db.find('orders', o => o.code === req.params.code);
  if (!order) return res.status(404).json({ error: 'Não encontrado' });
  const img = await QRCode.toBuffer(order.cardLink, { width: 300, errorCorrectionLevel: 'H', color: { dark: '#1E1B4B', light: '#ffffff' } });
  res.set('Content-Type', 'image/png');
  res.send(img);
});

// SPA fallback
app.get('/card/:code', (_, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));
app.get('*',           (_, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

app.listen(PORT, () => {
  console.log(`\n🎁  Vella Gift Card rodando em ${BASE_URL}`);
  console.log(`    Admin: diogo.hhonda@gmail.com / Diogo9968\n`);
});
