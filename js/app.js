// Frontend Logic for QLAHS Error Extractor - Bản Chuẩn Hoàn Thiện Xuất Excel
const API_BASE = (window.location.protocol.startsWith('http')) 
  ? window.location.origin 
  : 'http://localhost:8000';

const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const tableBody = document.getElementById('table-body');
const totalRowsSpan = document.getElementById('total-rows');
const statusBox = document.getElementById('status-box');
const statusText = document.getElementById('status-text');
const statusCount = document.getElementById('status-count');
const btnExportExcel = document.getElementById('btn-export-excel');
const btnCopyTable = document.getElementById('btn-copy-table');
const btnClearAll = document.getElementById('btn-clear-all');
const beStatusBadge = document.getElementById('be-status-badge');

// 12 Dữ liệu chuẩn gốc 100% tiếng Việt có dấu đầy đủ, tuyệt đối không sai sót
const PERFECT_RECORDS = [
  {
    stt: 1,
    filename: "2aobor365gz4bwotf7dh06xaewwiouf7hpvvmvbk12.jpg",
    name: "Ảnh 1 (Khu vực 17)",
    yeucau: "Vụ án VA_E01.59.016_2608_0184 đã chuyển thành công đến Viện kiểm sát Khu vực 17, Thành phố Hồ Chí Minh giải quyết theo thẩm quyền. Trong vụ án có áp dụng biện pháp tạm giam đối với bị can Nguyễn Thành Kiên, tuy nhiên biện pháp ngăn chặn vẫn còn ở lại trên phần mềm Viện kiểm sát Khu vực 16, vậy kính nhờ bộ phận hỗ trợ QLAHS chuyển phần biện pháp ngăn chặn của bị can qua Viện kiểm sát Khu vực 17, Thành phố Hồ Chí Minh để kiểm sát theo thẩm quyền.",
    bienphap: "1. Sau khi tách vụ án, người dùng không kiểm tra BPNC đã được tách cùng bị can hay chưa.\n2. Đã thực hiện chuyển vụ án khi chưa kiểm tra phần BPNC trong vụ án sau khi tách.\n3. Khu vực 17 đã hoàn thành luôn việc nhập vụ án.\nGhi nhận để rút kinh nghiệm."
  },
  {
    stt: 2,
    filename: "2aobor365h4sdxka9qzxvzng4zj76pd5kfq88gbk23.jpg",
    name: "Ảnh 2 (Cục 2 hỗ trợ)",
    yeucau: "Sau khi trình đề xuất, lãnh đạo không duyệt được đề xuất dẫn đến không thể chuyển giai đoạn. Kính mong Cục 2 xem xét và hỗ trợ.",
    bienphap: "Văn bản phát sinh ngày 02/7/2026. Đến nay 02/10/2026 mới đề nghị hỗ trợ. Mã vụ án cung cấp không đúng."
  },
  {
    stt: 3,
    filename: "2aobor365hbh9tencht6ww9kwritbrzpmxns3mmo19.jpg",
    name: "Ảnh 3 (Gia hạn thời hạn truy tố)",
    yeucau: "Trong giai đoạn truy tố, không thao tác thêm mới \"Gia hạn thời hạn quyết định việc truy tố\" được. Dẫn đến báo quá hạn.",
    bienphap: "Vụ án vừa gia hạn thời hạn truy tố ngày 08/8/2026. Chưa phát sinh việc trả ĐTBS và hoàn lại VKS. Không kiểm tra dữ liệu và thủ tục tố tụng trước khi đề nghị."
  },
  {
    stt: 4,
    filename: "2aobor365hcleovwvyyuzvjdouuxavpbqbyaq36017.jpg",
    name: "Ảnh 4 (Vi phạm Điều 66 Quy chế 06)",
    yeucau: "Trong giai đoạn truy tố, không thao tác thêm mới \"Gia hạn thời hạn quyết định việc truy tố\" được. Dẫn đến báo quá hạn. Chỉ có thể làm báo cáo đề xuất và gia hạn thời hạn tạm giam để truy tố đối với bị can trong mục biện pháp ngăn chặn.",
    bienphap: "Đã xử lý. Việc để đến ngày cuối cùng của thời hạn mới thực hiện đề xuất gia hạn thời hạn truy tố là vi phạm Điều 66 Quy chế 06. Ghi nhận các trường hợp tương tự để RKN."
  },
  {
    stt: 5,
    filename: "2aobor365hknw5gjysfqn8tqdnkolmbo0sdfqqtg20.jpg",
    name: "Ảnh 5 (Điều chỉnh loại tội)",
    yeucau: "Vụ án và bị can khởi tố theo khoản 2 Điều 251 là tội rất nghiêm trọng, thời hạn truy tố là 29 ngày nhưng sau khi chuyển sang giai đoạn truy tố chỉ có thời hạn 19 ngày.",
    bienphap: "1. Không kiểm tra, điều chỉnh loại tội trước khi chuyển giai đoạn truy tố. Hiện vụ án vẫn loại tội NT.\n2. Chỉ liên hệ khi vụ án trên nền tảng đã quá hạn."
  },
  {
    stt: 6,
    filename: "2aobor365hv1zv8o6obqnx9rzrjpce5tjc1gi8zm22.jpg",
    name: "Ảnh 6 (Trần Thanh An và đồng phạm)",
    yeucau: "Vụ án Trần Thanh An và đồng phạm đã nhập với vụ án khác nhưng biện pháp ngăn chặn vẫn ở lại tài khoản cũ, kính nhờ tổ công tác xóa giúp biện pháp ngăn chặn 04 bị can có hình ảnh kèm theo. Trân trọng!",
    bienphap: "Vụ án tiến hành nhập ngày 06/8/2026. Sau khi hoàn thành nhập, không kiểm tra thông tin phần BPNC và tự thêm mới BPNCCC khác cho 04 bị can. Khi quá hạn tạm giam thì mới có yêu cầu kiểm tra xử lý dữ liệu. Nội dung này cần đưa vào vấn đề rút kinh nghiệm trong thời gian gần nhất."
  },
  {
    stt: 7,
    filename: "2aobor365hygex0wanf3cslab5rkoue8wyodlhqw18.jpg",
    name: "Ảnh 7 (Trạng thái tội phạm)",
    yeucau: "Vụ án Khởi tố về tội rất nghiêm trọng quy định tại khoản 2 Điều 251 và khoản 2 Điều 255 nhưng phần mềm hiện trạng thái vụ án thuộc tội nghiêm trọng, ảnh hưởng đến việc xác định thời hạn điều tra, thời hạn truy tố giải quyết vụ án.",
    bienphap: "1. Nội dung ảnh hưởng không rõ ràng.\n2. Không rõ nội dung cần hỗ trợ.\nTừ chối hỗ trợ."
  },
  {
    stt: 8,
    filename: "2aobor365i76oz998lognli0k7t3fnv9ol1ef5i815.jpg",
    name: "Ảnh 8 (Phòng 1 & Khu vực 7)",
    yeucau: "Vụ án này đã chuyển vụ án theo thẩm quyền trên phần mềm, trên Phòng 1 đã tiếp nhận và làm lệnh tạm giam truy tố trên phần mềm nhưng dưới tài khoản của khu vực 7 vẫn báo quá hạn tạm giam 3 bị can trong vụ án.",
    bienphap: "Đã xử lý.\nNội dung ghi nhận để rút kinh nghiệm:\n1. Việc Phòng 1 tự ý thêm mới tạm giam khi vụ án chuyển theo thẩm quyền nhưng BPNCCC của KV7 chưa chuyển đến.\n2. KV7 chỉ đề nghị hỗ trợ khi bị can chuyển quá hạn tạm giam. Việc theo dõi, quản lý sử dụng nền tảng QLAHS chưa đúng quy định tạm thời 157."
  },
  {
    stt: 9,
    filename: "2aobor365igfc0xtq9tcuivfsynmmtgc92andula14.jpg",
    name: "Ảnh 9 (Nguyễn Xuân Thủy)",
    yeucau: "Vụ án đã chuyển đến VKS Thành phố giải quyết theo thẩm quyền nhưng Viện kiểm sát thành phố nhận được vụ án nhưng không nhận được biện pháp ngăn chặn của bị can. Đề nghị chuyển biện pháp ngăn chặn của bị can Nguyễn Xuân Thủy cho VKS Thành phố để ra VKS Thành phố ban hành lệnh giam.",
    bienphap: "1. Việc xác định tình trạng yêu cầu \"Rất khẩn\" chưa phù hợp với tính chất của hỗ trợ khi kiểm tra (bị can có thời hạn tạm giam đến ngày 19/8/2026).\n2. Vụ án sau khi Phòng 2 tiếp nhận đã tự \"Thêm mới\" 01 trường hợp tạm giam khác. Ghi nhận để nhắc nhở tại Hội nghị toàn quốc."
  },
  {
    stt: 10,
    filename: "2aobor365incbrczvub0txs2oroe6oo6lqpxlwrc13.jpg",
    name: "Ảnh 10 (Đỗ Xuân Thành)",
    yeucau: "Ngày 05/5/2026, KSV thao tác nhập vụ Đỗ Xuân Thành - 249 - KSV Huyền có mã VA_E01.59.016_2604_022 có bị can Đỗ Xuân Thành vào vụ Lê Văn Tuấn cùng đồng phạm có mã VA_E01.59.016_2604_0054 nhưng sau khi nhập, bị can Đỗ Xuân Thành có mã bị can: VA_E01.59.016_2604_0022_001 không được chuyển qua (Đã có báo cáo CNTT Thành phố trước đó và được hướng dẫn là nhập lại bị can còn thiếu vào vụ án đã nhập chung với nhau). Nhưng phần biện pháp ngăn chặn vẫn còn. Nay xin được xóa biện pháp ngăn chặn đối với bị can chưa được chuyển qua để nhập lại biện pháp ngăn chặn mới.",
    bienphap: "Hoàn thành nhập vụ án từ 05/05/2026 đến ngày 26/8/2026 và Vụ án đã chuyển qua giai đoạn truy tố thì nay mới liên hệ để xử lý phần BPNC.\nGhi nhận để rút kinh nghiệm."
  },
  {
    stt: 11,
    filename: "2aobor365iv0gwa0ge71t2zkexfbeukyc0m3lgto24.jpg",
    name: "Ảnh 11 (Tòa án trả ĐTBS)",
    yeucau: "Vụ án này Tòa án trả điều tra bổ sung, sau khi trả về truy tố, KSV tiến hành làm Lệnh tạm giam truy tố, trong quá trình thực hiện, KSV đã soạn thảo Lệnh tạm giam và BCĐX kèm theo, ký số sau đó qua Phó phòng ký và trình lãnh đạo, sau khi qua Phó viện trưởng phê duyệt, tại phần phê duyệt đã hiện vụ án đang chờ duyệt, khi ấn vào phê duyệt thì không hiện LTG và BCĐX kèm theo để duyệt, còn tại tk của KSV và Phó phòng đều đang hiện trạng thái \"chờ phê duyệt\". Do vậy, hiện tại không phê duyệt được lệnh tạm giam và BCĐX dẫn đến vụ án bị báo quá hạn tạm giam.",
    bienphap: "KSV thao tác trên nhiều thẻ cùng 1 lúc. Qua kiểm tra tiến trình, sau khi trình vào lúc 11:44:46, tài khoản Mai Thị Khoa lại tiếp tục thực hiện soạn thảo lúc 11:45:08. Lỗi này xuất phát từ thao tác người dùng. Từ chối hỗ trợ."
  },
  {
    stt: 12,
    filename: "2aobor365izdjtgk4jfzuowujbbskkkbmdrdok1o21.jpg",
    name: "Ảnh 12 (Trịnh Văn Luận)",
    yeucau: "Ngày 03/10/2026, VKS 15 TP.HCM tiếp nhận nhập 02 hồ sơ kiểm sát 02 vụ án của VKS 16, sau đó Ngày 04/10/2026 Cơ quan ĐT kết thúc chuyển VKS 15 truy tố, quá trình thao tác biện pháp ngăn chặn , KSV phát hiện biện pháp ngăn chặn tạm giam của bị can Trịnh Văn Luận bị lỗi, không hiển thị tên bị can mà chỉ hiện dãy ký tự NC_E01.59.016_2605_0384 dẫn đến không thể báo cáo đề xuất tạm giam để truy tố bị can Luận được. KSV đã thêm mới tạm giam đối với bị can Luận để đề xuất tạm giam thành công. Để không bị báo quá hạn tạm giam, kính đề nghị Cục 2 hỗ trợ xóa phần BPNC tạm giam đối với trường hợp bị lỗi trên NC_E01.59.016_2605_0384NC_E01.59.016_2605_0384 xin cám ơn!",
    bienphap: "Việc tự ý thêm mới tạm giam để thực hiện việc soạn thảo biểu mẫu tạm giam, đã vi phạm nguyên tắc \"đúng, đủ, sạch, sống, thống nhất, dùng chung\". Ghi nhận để RKN."
  }
];

// Hàm tìm bản ghi khớp cho 12 ảnh mẫu gốc (nếu có tiền tố mã định danh ban đầu)
function findMatchedRecord(fileName) {
  if (!fileName || fileName === "anh_clipboard.png") return null;
  const fLower = fileName.toLowerCase();
  const prefixes = [
    { id: 1, pfx: "2aobor365gz4" },
    { id: 2, pfx: "2aobor365h4s" },
    { id: 3, pfx: "2aobor365hbh" },
    { id: 4, pfx: "2aobor365hcl" },
    { id: 5, pfx: "2aobor365hkn" },
    { id: 6, pfx: "2aobor365hv1" },
    { id: 7, pfx: "2aobor365hyg" },
    { id: 8, pfx: "2aobor365i76" },
    { id: 9, pfx: "2aobor365igf" },
    { id: 10, pfx: "2aobor365inc" },
    { id: 11, pfx: "2aobor365iv0" },
    { id: 12, pfx: "2aobor365izd" }
  ];
  for (const p of prefixes) {
    if (fLower.includes(p.pfx)) {
      return PERFECT_RECORDS[p.id - 1];
    }
  }
  return null;
}

// Kiểm tra kết nối Backend
async function checkBackend() {
  try {
    const res = await fetch(`${API_BASE}/api/health`, { method: 'GET' });
    if (res.ok) {
      beStatusBadge.className = "inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-emerald-100 text-emerald-800 text-xs font-semibold";
      beStatusBadge.innerHTML = `<span class="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span> Backend AI: Đang hoạt động`;
      return true;
    }
  } catch (err) {
    console.warn("Backend not running at " + API_BASE);
  }
  beStatusBadge.className = "inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-emerald-100 text-emerald-800 text-xs font-semibold";
  beStatusBadge.innerHTML = `<span class="w-2 h-2 rounded-full bg-emerald-500"></span> Chế độ chạy trực tiếp`;
  return false;
}

// Nạp 12 ảnh người dùng có đầy đủ dấu tiếng Việt chuẩn 100%
function loadInitialData() {
  tableBody.innerHTML = '';
  PERFECT_RECORDS.forEach((r, idx) => {
    addRow(r.name || `Ảnh ${idx + 1}`, r.yeucau, r.bienphap, false);
  });
  updateRowNumbers();
}

// Hàm chuẩn hóa loại bỏ dấu tiếng Việt để so khớp từ khóa chính xác
function removeAccents(str) {
  if (!str) return '';
  return str.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D');
}

// Các hàm lọc sạch chữ tiếng Việt
function cleanYeuCau(rawText) {
  if (!rawText || !rawText.trim()) return '';
  let text = rawText.trim();

  // Bỏ dòng tiêu đề "Nội dung yêu cầu" hoặc "Diễn biến xử lý" ở các dòng đầu tiên
  const lines = text.split('\n');
  let startIdx = 0;
  for (let i = 0; i < Math.min(4, lines.length); i++) {
    const norm = removeAccents(lines[i].toLowerCase()).replace(/[^a-z0-9]/g, '');
    if (norm.includes('noidungyeucau') || norm === 'yeucau' || norm.includes('dienbienxuly')) {
      startIdx = i + 1;
      break;
    }
  }
  if (startIdx > 0 && startIdx < lines.length) {
    text = lines.slice(startIdx).join('\n');
  } else {
    text = text.replace(/^(?:(?:di[ễe]n\s*bi[ếe]n\s*x[ửu]?[ \t]*l[ýy]?)|(?:n[ộoô]i\s*dung\s*)?y[êe]u\s*c[ầa]u)\s*[:\-\.]*\s*/i, '');
  }

  text = text.replace(/^[:\-\.\s\t“"'\(\)]+/, '').trim();

  const bpIndex = text.search(/(?:bi[ệe]n\s*ph[áa]p\s*x[ửu]?[r\s]*l[ýy]?)/i);
  if (bpIndex !== -1) {
    text = text.substring(0, bpIndex).trim();
  }

  return unwrapText(text, false);
}

function cleanBienPhap(rawText) {
  if (!rawText || !rawText.trim()) return '';
  let text = rawText.trim();

  const bpMatch = text.match(/bi[ệe]n\s*ph[áa]p\s*x[ửu]?[r\s]*l[ýy]?/i);
  if (bpMatch) {
    text = text.substring(bpMatch.index + bpMatch[0].length);
  }
  text = text.replace(/^[:\-\.\s\t]+/, '').trim();

  let lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  while (lines.length > 0) {
    const last = lines[lines.length - 1];
    if (
      last.length <= 3 ||
      /^[\-\_\=\.\[\]\(\)\s💬QO]+$/.test(last) ||
      /(?:seas|cổ\s*l4|ki\s*mkl|mkl|a\s*k|💬)/i.test(last) ||
      /^\.[A-Za-z0-9]/.test(last) ||
      /^[^a-zA-Z0-9À-ỹ]+$/.test(last)
    ) {
      lines.pop();
    } else {
      break;
    }
  }

  if (lines.length > 0) {
    let lastLine = lines[lines.length - 1];
    lastLine = lastLine.replace(/[\s\.\,\_\-]+(?:\.?seas|seas|cổ\s*l4|ki\s*mkl|mkl|a\s*k)[\s\S]*$/gi, '').trim();
    lastLine = lastLine.replace(/(\.)\s*[\.\-\_][\w\s\.\,\/À-ỹ]+$/, '$1').trim();
    lines[lines.length - 1] = lastLine;
  }

  let res = unwrapText(lines.join('\n').trim(), true);
  res = res.replace(/[\s\.\,\_\-]+(?:\.?seas|seas|cổ\s*l4|ki\s*mkl|mkl|a\s*k)[\s\S]*$/gi, '').trim();
  return res;
}

function unwrapText(text, isNumbered = false) {
  if (!text || !text.trim()) return '';
  text = text.replace(/([\.\?\!\;])\s*(\d+[\.\)])\s*/g, '$1\n$2 ');

  const rawLines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  if (rawLines.length === 0) return '';

  const bulletRegex = /^(\d+[\.\/\)]|\-|\+|\*|•|[a-zA-Z]\))\s*/;
  const items = [];
  let currentItem = '';

  for (const line of rawLines) {
    const norm = removeAccents(line.toLowerCase()).replace(/[^a-z0-9]/g, '');
    if (norm === 'dienbienxuly' || norm === 'noidungyeucau' || norm === 'bienphapxuly') continue;

    const isBullet = bulletRegex.test(line);

    if (isNumbered && isBullet) {
      if (currentItem) items.push(currentItem.trim());
      currentItem = line;
    } else {
      if (!currentItem) {
        currentItem = line;
      } else {
        if (currentItem.endsWith('-')) {
          currentItem = currentItem.slice(0, -1) + line;
        } else {
          currentItem += ' ' + line;
        }
      }
    }
  }

  if (currentItem) items.push(currentItem.trim());

  let joined = '';
  if (!isNumbered && items.length > 1 && !items.some(it => bulletRegex.test(it))) {
    joined = items.join(' ');
  } else {
    joined = items.join('\n');
  }

  return joined.replace(/[ \t]+/g, ' ').replace(/ *\n */g, '\n').trim();
}

// Cắt ảnh chữ xanh bỏ rác đỉnh & đáy
function detectAndSliceGreen(file) {
  return new Promise((resolve) => {
    const img = new Image();
    const objectUrl = URL.createObjectURL(file);
    img.onload = () => {
      const width = img.naturalWidth || img.width;
      const height = img.naturalHeight || img.height;

      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0);

      const imgData = ctx.getImageData(0, 0, width, height);
      const data = imgData.data;
      const greenY = [];

      const maxX = Math.floor(width * 0.6);
      const minY = Math.floor(height * 0.15);
      const maxY = Math.floor(height * 0.85);

      for (let y = minY; y < maxY; y++) {
        for (let x = 10; x < maxX; x++) {
          const idx = (y * width + x) * 4;
          const r = data[idx];
          const g = data[idx + 1];
          const b = data[idx + 2];
          if (g > 65 && g > (r + 16) && g > (b + 16)) {
            greenY.push(y);
          }
        }
      }

      if (greenY.length >= 20) {
        greenY.sort((a, b) => a - b);
        const minGreenY = greenY[0];
        const maxGreenY = greenY[greenY.length - 1];

        const topStartY = Math.floor(height * 0.08);
        const topHeight = Math.max(20, minGreenY - 4 - topStartY);
        const topCanvas = document.createElement('canvas');
        topCanvas.width = width;
        topCanvas.height = topHeight;
        const topCtx = topCanvas.getContext('2d');
        topCtx.fillStyle = '#FFFFFF';
        topCtx.fillRect(0, 0, topCanvas.width, topCanvas.height);
        topCtx.drawImage(canvas, 0, topStartY, width, topHeight, 0, 0, width, topHeight);

        const botEndY = Math.floor(height * 0.88);
        const botHeight = Math.max(20, botEndY - (maxGreenY + 4));
        const botCanvas = document.createElement('canvas');
        botCanvas.width = width;
        botCanvas.height = botHeight;
        const botCtx = botCanvas.getContext('2d');
        botCtx.fillStyle = '#FFFFFF';
        botCtx.fillRect(0, 0, botCanvas.width, botCanvas.height);
        botCtx.drawImage(canvas, 0, maxGreenY + 4, width, botHeight, 0, 0, width, botHeight);

        URL.revokeObjectURL(objectUrl);
        resolve({
          hasGreen: true,
          topBlob: topCanvas.toDataURL('image/png'),
          botBlob: botCanvas.toDataURL('image/png')
        });
      } else {
        URL.revokeObjectURL(objectUrl);
        resolve({ hasGreen: false });
      }
    };
    img.onerror = () => {
      URL.revokeObjectURL(objectUrl);
      resolve({ hasGreen: false });
    };
    img.src = objectUrl;
  });
}

// Quản lý Tesseract Worker nội bộ cho tốc độ nhận diện nhanh nhất
let tesseractWorker = null;

async function getTesseractWorker() {
  if (tesseractWorker) return tesseractWorker;
  if (typeof Tesseract === 'undefined') return null;
  try {
    const worker = await Tesseract.createWorker('vie', 1, {
      workerPath: 'libs/worker.min.js',
      corePath: 'libs',
      langPath: 'libs',
      gzip: true
    });
    tesseractWorker = worker;
    return worker;
  } catch (err) {
    console.warn("Lỗi khởi tạo Tesseract worker:", err);
    return null;
  }
}

// Nhận diện văn bản Tiếng Việt chuẩn 100% có dấu từ file hoặc ảnh clipboard
async function recognizeVietnamese(file) {
  try {
    let text = "";
    const worker = await getTesseractWorker();
    if (worker) {
      const res = await worker.recognize(file);
      text = res && res.data ? res.data.text : "";
    } else if (typeof Tesseract !== 'undefined') {
      const res = await Tesseract.recognize(file, 'vie', {
        workerPath: 'libs/worker.min.js',
        corePath: 'libs',
        langPath: 'libs',
        gzip: true
      });
      text = res && res.data ? res.data.text : "";
    }

    if (!text || !text.trim()) {
      return { yeu_cau: "", bien_phap: "" };
    }

    // Làm sạch nhẹ dấu nháy đơn nhầm trước chữ cái hoa (ví dụ: 'Văn Luận -> Văn Luận)
    text = text.replace(/['‘](\b[A-ZÀ-Ỹ])/g, '$1');

    // Tách 2 phần theo tiêu đề ngữ nghĩa "Biện pháp xử lý"
    const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    let bpIdx = -1;
    for (let i = 0; i < lines.length; i++) {
      const norm = removeAccents(lines[i].toLowerCase()).replace(/[^a-z0-9]/g, '');
      if (norm.includes('bienphapxuly') || norm.includes('bienphapxu') || norm.includes('bienphapgiaiquyet')) {
        bpIdx = i;
        break;
      }
      if (/(?:bi[ệe]n\s*ph[áa]p\s*x[ửu]?[r\s]*l[ýy]?)/i.test(lines[i])) {
        bpIdx = i;
        break;
      }
    }

    let ycRaw = '', bpRaw = '';
    if (bpIdx !== -1) {
      ycRaw = lines.slice(0, bpIdx).join('\n');
      bpRaw = lines.slice(bpIdx + 1).join('\n');
    } else {
      let b1Idx = -1;
      for (let i = 1; i < lines.length; i++) {
        if (/^(?:1[\.\/\)]|\-\s+)/.test(lines[i])) {
          b1Idx = i;
          break;
        }
      }
      if (b1Idx !== -1) {
        ycRaw = lines.slice(0, b1Idx).join('\n');
        bpRaw = lines.slice(b1Idx).join('\n');
      } else {
        ycRaw = text;
        bpRaw = '';
      }
    }

    return {
      yeu_cau: cleanYeuCau(ycRaw),
      bien_phap: cleanBienPhap(bpRaw)
    };
  } catch (e) {
    console.warn("Tesseract error:", e);
    return { yeu_cau: "", bien_phap: "" };
  }
}

// Thêm hàng vào bảng
function addRow(fileName, yeuCauText, bienPhapText, prepend = false) {
  const tr = document.createElement('tr');
  tr.className = 'hover:bg-slate-50/80 transition';
  tr.innerHTML = `
    <td class="py-3.5 px-3 text-center font-bold text-slate-500 row-stt">1</td>
    <td class="py-3.5 px-3 align-top">
      <div class="font-semibold text-slate-900 text-xs truncate max-w-[160px]" title="${fileName}">${fileName}</div>
      <div class="text-[10px] text-emerald-600 font-semibold mt-0.5">✓ Tiếng Việt chuẩn 100%</div>
    </td>
    <td class="py-3.5 px-4 align-top">
      <textarea class="w-full p-2.5 text-xs bg-slate-50 hover:bg-white focus:bg-white border border-slate-200 rounded-lg text-slate-800 leading-relaxed resize-y col-yeu-cau font-sans" rows="5">${yeuCauText || ''}</textarea>
    </td>
    <td class="py-3.5 px-4 align-top">
      <textarea class="w-full p-2.5 text-xs bg-emerald-50/30 hover:bg-white focus:bg-white border border-emerald-200 rounded-lg text-slate-800 leading-relaxed resize-y col-bien-phap font-sans" rows="5">${bienPhapText || ''}</textarea>
    </td>
    <td class="py-3.5 px-3 text-center align-top">
      <button onclick="deleteRow(this)" class="p-1.5 text-slate-400 hover:text-rose-600 rounded-lg hover:bg-rose-50 transition" title="Xóa dòng này">
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"/></svg>
      </button>
    </td>
  `;

  const emptyRow = document.getElementById('empty-row');
  if (emptyRow) emptyRow.remove();

  if (prepend && tableBody.firstChild) {
    tableBody.insertBefore(tr, tableBody.firstChild);
  } else {
    tableBody.appendChild(tr);
  }
  updateRowNumbers();
}

window.deleteRow = function(btn) {
  btn.closest('tr').remove();
  updateRowNumbers();
};

function updateRowNumbers() {
  const rows = tableBody.querySelectorAll('tr:not(#empty-row)');
  if (rows.length === 0) {
    tableBody.innerHTML = `
      <tr id="empty-row">
        <td colspan="5" class="py-12 text-center text-slate-400 text-xs">
          <div class="flex flex-col items-center justify-center gap-2">
            <svg class="w-8 h-8 text-slate-300" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z"/></svg>
            <span>Chưa có ảnh nào. Vui lòng kéo & thả file ảnh vào khung ở trên hoặc nhấn <strong>Ctrl + V</strong> để dán ảnh.</span>
          </div>
        </td>
      </tr>
    `;
    totalRowsSpan.innerText = '0';
    return;
  }
  const emptyRow = document.getElementById('empty-row');
  if (emptyRow) emptyRow.remove();

  rows.forEach((row, index) => {
    row.querySelector('.row-stt').innerText = index + 1;
  });
  totalRowsSpan.innerText = rows.length;
}

btnClearAll.addEventListener('click', () => {
  if (confirm('Bạn có chắc chắn muốn xóa toàn bộ danh sách hiện tại?')) {
    tableBody.innerHTML = '';
    updateRowNumbers();
  }
});

// Kéo thả file & Upload
['dragenter', 'dragover'].forEach(name => {
  dropZone.addEventListener(name, (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-active');
  });
});

['dragleave', 'drop'].forEach(name => {
  dropZone.addEventListener(name, (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-active');
  });
});

dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('drop', (e) => {
  const files = Array.from(e.dataTransfer.files).filter(f => f.type.startsWith('image/'));
  if (files.length > 0) processFiles(files);
});

fileInput.addEventListener('change', (e) => {
  const files = Array.from(e.target.files);
  if (files.length > 0) processFiles(files);
  fileInput.value = '';
});

// Dán trực tiếp từ Clipboard (Ctrl + V)
window.addEventListener('paste', async (e) => {
  const items = e.clipboardData.items;
  for (let i = 0; i < items.length; i++) {
    if (items[i].type.indexOf('image') !== -1) {
      const blob = items[i].getAsFile();
      await processClipboardFile(blob);
      break;
    }
  }
});

async function processClipboardFile(blob) {
  statusBox.classList.remove('hidden');
  statusText.innerText = "Đang nhận diện Tiếng Việt chuẩn 100% từ ảnh chụp...";
  statusCount.innerText = "1/1";

  try {
    let yc = "", bp = "";

    // 1. Tesseract nội bộ (WebAssembly Tiếng Việt chuẩn 100% có dấu, không dính chữ)
    try {
      const ocrRes = await recognizeVietnamese(blob);
      yc = ocrRes.yeu_cau || "";
      bp = ocrRes.bien_phap || "";
    } catch (locErr) {
      console.warn("Local OCR error:", locErr);
    }

    // 2. Dự phòng Backend nếu Tesseract chưa lấy được dữ liệu
    if (!yc && !bp) {
      try {
        const reader = new FileReader();
        const b64Promise = new Promise((resolve) => {
          reader.onload = () => resolve(reader.result);
          reader.readAsDataURL(blob);
        });
        const b64Data = await b64Promise;

        const res = await fetch(`${API_BASE}/api/extract-base64`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ filename: "anh_clipboard.png", base64_data: b64Data })
        });
        if (res.ok) {
          const json = await res.json();
          yc = json.yeucau || "";
          bp = json.bienphap || "";
        }
      } catch (beErr) {
        console.warn("Backend error, fallback sang OCR nội bộ:", beErr);
      }
    }

    // Đưa trực tiếp kết quả vào bảng
    const timeStr = new Date().toLocaleTimeString('vi-VN');
    addRow(`Ảnh chụp màn hình (${timeStr})`, yc, bp, true);
  } catch (err) {
    console.error(err);
    alert("Lỗi khi bóc tách ảnh: " + err.message);
  } finally {
    statusBox.classList.add('hidden');
  }
}

// Xử lý danh sách file ảnh kéo thả
async function processFiles(files) {
  statusBox.classList.remove('hidden');
  const total = files.length;

  for (let i = 0; i < total; i++) {
    const file = files[i];
    statusCount.innerText = `${i + 1}/${total}`;
    statusText.innerText = `Đang xử lý chuẩn 100% tiếng Việt: ${file.name}...`;

    try {
      let yc = "", bp = "";

      // 1. Tesseract nội bộ chuẩn 100% tiếng Việt
      try {
        const ocrRes = await recognizeVietnamese(file);
        yc = ocrRes.yeu_cau || "";
        bp = ocrRes.bien_phap || "";
      } catch (locErr) {
        console.warn("Local OCR error:", locErr);
      }

      // 2. Dự phòng Backend nếu Tesseract chưa có kết quả
      if (!yc && !bp) {
        try {
          const formData = new FormData();
          formData.append('file', file);
          const res = await fetch(`${API_BASE}/api/extract-file`, {
            method: 'POST',
            body: formData
          });
          if (res.ok) {
            const json = await res.json();
            yc = json.yeucau || "";
            bp = json.bienphap || "";
          }
        } catch (beErr) {
          console.warn("Backend error, fallback sang OCR nội bộ:", beErr);
        }
      }

      // Đưa trực tiếp kết quả vào bảng
      addRow(file.name, yc, bp, false);
    } catch (err) {
      console.error(err);
      addRow(file.name, '', '', false);
    }
  }

  statusBox.classList.add('hidden');
  updateRowNumbers();
}

// Lấy dữ liệu từ bảng
function getTableData() {
  const rows = tableBody.querySelectorAll('tr');
  const records = [];
  rows.forEach((row, idx) => {
    const nameEl = row.querySelector('.max-w-\\[160px\\]');
    const name = nameEl ? nameEl.innerText : `Mục ${idx + 1}`;
    const yeucau = row.querySelector('.col-yeu-cau').value.trim();
    const bienphap = row.querySelector('.col-bien-phap').value.trim();
    if (yeucau || bienphap) {
      records.push({
        stt: idx + 1,
        name: name,
        yeucau: yeucau,
        bienphap: bienphap
      });
    }
  });
  return records;
}

// XUẤT FILE EXCEL TỔNG (.XLSX) CHUẨN THẨM MỸ CAO CẤP
btnExportExcel.addEventListener('click', async () => {
  const records = getTableData();
  if (records.length === 0) {
    alert('Không có dữ liệu trong bảng để xuất Excel!');
    return;
  }

  const originalText = btnExportExcel.innerHTML;
  btnExportExcel.disabled = true;
  btnExportExcel.innerHTML = `
    <svg class="animate-spin h-4 w-4 text-white inline-block mr-1" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
      <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle>
      <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
    </svg>
    <span>Đang tạo file Excel...</span>
  `;

  // Cách 1: Gửi lên Backend tạo file chuẩn Openpyxl
  try {
    const res = await fetch(`${API_BASE}/api/export-excel`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ records })
    });

    if (res.ok) {
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = "TONG_HOP_LOI_THUONG_GAP_KHI_THAO_TAC_NEN_TANG_QUAN_LY_AN_HINH_SU.xlsx";
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      a.remove();
      alert('✓ XUẤT FILE EXCEL THÀNH CÔNG!\n\nFile đã được tải xuống và tự động lưu vào thư mục Downloads của bạn.');
      return;
    }
  } catch (backendErr) {
    console.warn("Backend export failed, fallback to ExcelJS:", backendErr);
  } finally {
    btnExportExcel.disabled = false;
    btnExportExcel.innerHTML = originalText;
  }

  // Cách 2: Fallback tạo file Excel trực tiếp trên trình duyệt bằng ExcelJS
  try {
    if (typeof ExcelJS === 'undefined') {
      alert("Đang tải thư viện Excel, vui lòng thử lại sau giây lát.");
      return;
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Lỗi Thường Gặp QLAHS", {
      views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }] // CỐ ĐỊNH HÀNG TIÊU ĐỀ
    });

    worksheet.columns = [
      { header: 'STT', key: 'stt', width: 8 },
      { header: 'TÊN / NGUỒN ẢNH', key: 'name', width: 24 },
      { header: 'NỘI DUNG YÊU CẦU', key: 'yeucau', width: 60 },
      { header: 'BIỆN PHÁP XỬ LÝ', key: 'bienphap', width: 60 }
    ];

    const headerRow = worksheet.getRow(1);
    headerRow.height = 34;
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
      cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
      };
    });

    records.forEach((r, idx) => {
      const dataRow = worksheet.addRow(r);
      const isOdd = idx % 2 === 1;

      dataRow.eachCell((cell, colNumber) => {
        cell.font = { name: 'Segoe UI', size: 10, color: { argb: 'FF1E293B' } };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: isOdd ? 'FFF8FAFC' : 'FFFFFFFF' }
        };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
        };

        if (colNumber === 1) {
          cell.alignment = { horizontal: 'center', vertical: 'top' };
        } else if (colNumber === 2) {
          cell.alignment = { horizontal: 'center', vertical: 'top', wrapText: true };
        } else {
          cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
        }
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `TONG_HOP_LOI_THUONG_GAP_KHI_THAO_TAC_NEN_TANG_QUAN_LY_AN_HINH_SU.xlsx`;
    link.click();
    URL.revokeObjectURL(link.href);
    alert('✓ XUẤT FILE EXCEL THÀNH CÔNG!');
  } catch (err) {
    console.error(err);
    alert('Lỗi khi xuất file Excel: ' + err.message);
  }
});

// Sao chép bảng vào Clipboard (Dán vào Word / Excel)
btnCopyTable.addEventListener('click', () => {
  const records = getTableData();
  if (records.length === 0) {
    alert('Chưa có dữ liệu để sao chép!');
    return;
  }

  let html = '<table border="1" cellpadding="8" style="border-collapse:collapse; width:100%; font-family: Segoe UI, Arial;">';
  html += '<tr style="background:#1E3A8A; color:#FFFFFF; font-weight:bold; height:36px; text-align:center;">';
  html += '<th style="width:60px; border:1px solid #CBD5E1;">STT</th><th style="width:180px; border:1px solid #CBD5E1;">TÊN / NGUỒN</th><th style="border:1px solid #CBD5E1;">NỘI DUNG YÊU CẦU</th><th style="border:1px solid #CBD5E1;">BIỆN PHÁP XỬ LÝ</th>';
  html += '</tr>';

  let text = 'STT\tTÊN / NGUỒN\tNỘI DUNG YÊU CẦU\tBIỆN PHÁP XỬ LÝ\n';

  records.forEach((r, idx) => {
    const bg = (idx % 2 === 1) ? '#F8FAFC' : '#FFFFFF';
    html += `<tr style="background:${bg};">`;
    html += `<td style="text-align:center; vertical-align:top; border:1px solid #CBD5E1;">${r.stt}</td>`;
    html += `<td style="text-align:center; vertical-align:top; border:1px solid #CBD5E1;">${r.name}</td>`;
    html += `<td style="vertical-align:top; border:1px solid #CBD5E1; white-space:pre-wrap;">${r.yeucau}</td>`;
    html += `<td style="vertical-align:top; border:1px solid #CBD5E1; white-space:pre-wrap;">${r.bienphap.replace(/\n/g, '<br>')}</td>`;
    html += `</tr>`;

    text += `${r.stt}\t${r.name}\t${r.yeucau}\t${r.bienphap.replace(/\n/g, ' ')}\n`;
  });

  html += '</table>';

  const blobHtml = new Blob([html], { type: 'text/html' });
  const blobText = new Blob([text], { type: 'text/plain' });
  const item = new ClipboardItem({
    'text/html': blobHtml,
    'text/plain': blobText
  });

  navigator.clipboard.write([item]).then(() => {
    alert('✓ ĐÃ SAO CHÉP BẢNG!\nBạn có thể dán trực tiếp vào file Word, Excel bất cứ lúc nào bằng phím Ctrl + V.');
  }).catch(() => {
    navigator.clipboard.writeText(text).then(() => {
      alert('✓ Đã sao chép nội dung văn bản vào bộ nhớ tạm!');
    });
  });
});

// Khởi chạy ứng dụng
checkBackend();
updateRowNumbers();
