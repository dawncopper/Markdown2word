// ===== AI转Word助手 — 常量定义 =====

// ===== 中文传统字号 =====
const CHINESE_SIZES = [
    { n: '初号', p: 42 }, { n: '小初', p: 36 }, { n: '一号', p: 26 },
    { n: '小一', p: 24 }, { n: '二号', p: 22 }, { n: '小二', p: 18 },
    { n: '三号', p: 16 }, { n: '小三', p: 15 }, { n: '四号', p: 14 },
    { n: '小四', p: 12 }, { n: '五号', p: 10.5 }, { n: '小五', p: 9 }
];

/**
 * 将磅值(pt)转换为中文传统字号
 * @param {number} pt - 磅值
 * @returns {string} 中文字号或 "Xpt" 格式
 */
function ptToChinese(pt) {
    for (const s of CHINESE_SIZES) {
        if (Math.abs(s.p - pt) < 0.8) return s.n;
    }
    return `${pt}pt`;
}

/**
 * 将值转换为磅值
 * @param {string|number} val - 中文字号或数字
 * @returns {number} 磅值
 */
function chineseToPt(val) {
    if (!val) return 12;
    const num = parseFloat(val);
    if (!isNaN(num)) return num;
    for (const s of CHINESE_SIZES) {
        if (s.n === val) return s.p;
    }
    return 12;
}

// ===== 内置模板 =====
const DEFAULT_TEMPLATES = {
    official: {
        name: '公文模板',
        icon: '📄',
        desc: '党政机关公文格式 (GB/T 9704)',
        headingFont: 'SimHei',
        headingSize: 16,      // 三号
        bodyFont: 'FangSong',
        bodySize: 16,         // 三号
        heading2Font: 'KaiTi',
        heading2Size: 16,
        lineHeight: 28,        // 28磅固定值
        lineRule: 'exact',     // exact = 磅值, auto = 倍数
        alignment: 'justified',
        firstLineIndent: 2,    // 字符数
        indentTwips: 320,      // 2字符 ≈ 320twips
        margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        spacingBefore: 0,
        spacingAfter: 0
    },
    workplace: {
        name: '职场模板',
        icon: '💼',
        desc: '商务办公文档',
        headingFont: 'MicrosoftYaHei',
        headingSize: 15,       // 小三
        bodyFont: 'MicrosoftYaHei',
        bodySize: 12,          // 小四
        heading2Font: 'MicrosoftYaHei',
        heading2Size: 14,
        lineHeight: 1.5,
        lineRule: 'auto',
        alignment: 'left',
        firstLineIndent: 0,
        indentTwips: 0,
        margins: { top: 1270, right: 1270, bottom: 1270, left: 1270 },
        spacingBefore: 6,
        spacingAfter: 6
    },
    report: {
        name: '报告模板',
        icon: '📊',
        desc: '数据分析报告',
        headingFont: 'SimHei',
        headingSize: 15,       // 小三
        bodyFont: 'SimSun',
        bodySize: 12,          // 小四
        heading2Font: 'SimHei',
        heading2Size: 14,
        lineHeight: 1.5,
        lineRule: 'auto',
        alignment: 'left',
        firstLineIndent: 2,
        indentTwips: 320,
        margins: { top: 1270, right: 1270, bottom: 1270, left: 1270 },
        spacingBefore: 6,
        spacingAfter: 6
    }
};

const MODE_NAMES = {
    purify: '净化',
    rearrange: '重排',
    general: '通用'
};

// 主题配置
const THEME_CONFIG = {
    themes: ['light', 'dark', 'warm'],
    icons: { light: '🌙', dark: '☀️', warm: '📖' }
};

export {
    CHINESE_SIZES,
    ptToChinese,
    chineseToPt,
    DEFAULT_TEMPLATES,
    MODE_NAMES,
    THEME_CONFIG
};
