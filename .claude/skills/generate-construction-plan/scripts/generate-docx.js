#!/usr/bin/env node
/**
 * 施工组织设计 Word 文档生成器
 * 基于 docx npm 包创建格式规范的 .docx 文件
 *
 * 用法: node generate-docx.js --input <input.json> --output <output.docx>
 *       node generate-docx.js --stdin < output.json  # 从 stdin 读取
 *
 * 输入 JSON 格式:
 * {
 *   "projectName": "项目名称",
 *   "companyName": "编制单位",
 *   "date": "2026年6月",
 *   "chapters": [
 *     {
 *       "title": "章节标题",
 *       "level": 1,
 *       "content": [
 *         { "type": "paragraph", "text": "段落文字", "format": { "bold": false, "size": 24, "alignment": "both", "firstLineIndent": 2 } },
 *         { "type": "heading", "text": "子标题", "level": 2 },
 *         { "type": "list", "items": ["项1", "项2"], "listType": "bullet" },
 *         { "type": "table", "headers": ["列1", "列2"], "rows": [["A", "B"]] }
 *       ]
 *     }
 *   ]
 * }
 */

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType,
  PageBreak, TableOfContents, Header, Footer, PageNumber,
  NumberFormat, convertInchesToTwip, LevelFormat, TabStopPosition, TabStopType,
  ImageRun
} = require('docx');

// --- 字体与样式常量 ---
const FONT_ZH = '宋体';
const FONT_EN = 'Times New Roman';
const FONT_TITLE = '黑体';
const COLOR_TITLE = '1F3864';

// --- 中文字数估算 ---
// A4 标准页（含标题、表格）约容纳 500-600 汉字
const CHARS_PER_PAGE = 550;

// ========== 辅助函数 ==========

function getHeadingStyle(level) {
  const map = {
    1: HeadingLevel.HEADING_1,
    2: HeadingLevel.HEADING_2,
    3: HeadingLevel.HEADING_3,
    4: HeadingLevel.HEADING_4,
  };
  return map[level] || HeadingLevel.HEADING_2;
}

function createTextRun(text, format = {}) {
  const opts = {
    text: String(text),
    font: { name: FONT_EN, eastAsia: FONT_ZH },
    size: format.size || 24, // half-points: 24 = 12pt
    bold: format.bold || false,
    italics: format.italics || false,
    color: format.color || '000000',
  };
  if (format.font) opts.font = format.font;
  return new TextRun(opts);
}

function createParagraph(textOrRuns, options = {}) {
  const children = Array.isArray(textOrRuns)
    ? textOrRuns
    : [createTextRun(textOrRuns, options)];

  const pOpts = {
    children,
    spacing: { after: 120, line: 360 },
  };

  if (options.alignment) {
    const alignMap = {
      center: AlignmentType.CENTER,
      left: AlignmentType.LEFT,
      right: AlignmentType.RIGHT,
      both: AlignmentType.JUSTIFIED,
    };
    pOpts.alignment = alignMap[options.alignment] || AlignmentType.BOTH;
  }

  // 首行缩进（2字符 ≈ 480 twips）
  if (options.firstLineIndent) {
    pOpts.indent = { firstLine: options.firstLineIndent * 240 };
  }

  if (options.heading) {
    pOpts.heading = getHeadingStyle(options.heading);
  }

  if (options.pageBreak) {
    children.push(new PageBreak());
  }

  return new Paragraph(pOpts);
}

function createCoverPage(projectName, subtitle, companyName, dateStr) {
  const paragraphs = [];

  // 空行占位
  for (let i = 0; i < 6; i++) {
    paragraphs.push(new Paragraph({ children: [], spacing: { after: 0, line: 480 } }));
  }

  // 主标题
  paragraphs.push(createParagraph(projectName || '工程名称', {
    bold: true, size: 44, alignment: 'center', color: COLOR_TITLE,
    font: { name: FONT_EN, eastAsia: FONT_TITLE },
  }));

  paragraphs.push(new Paragraph({ children: [], spacing: { after: 200 } }));

  // 副标题
  paragraphs.push(createParagraph(subtitle || '施工组织设计', {
    bold: true, size: 36, alignment: 'center', color: COLOR_TITLE,
    font: { name: FONT_EN, eastAsia: FONT_TITLE },
  }));

  // 空行
  for (let i = 0; i < 8; i++) {
    paragraphs.push(new Paragraph({ children: [], spacing: { after: 0, line: 480 } }));
  }

  // 编制单位
  if (companyName) {
    paragraphs.push(createParagraph(`编制单位：${companyName}`, {
      size: 28, alignment: 'center',
    }));
    paragraphs.push(new Paragraph({ children: [], spacing: { after: 100 } }));
  }

  // 日期
  paragraphs.push(createParagraph(dateStr || new Date().toLocaleDateString('zh-CN'), {
    size: 28, alignment: 'center',
  }));

  // 分页
  paragraphs.push(new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } }));

  return paragraphs;
}

function createTableSection(headers, rows, caption, format = {}) {
  if (!headers || headers.length === 0) return [];

  const colCount = headers.length;

  // 创建表头行
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map(h => new TableCell({
      children: [createParagraph(String(h), { bold: true, size: 22, alignment: 'center' })],
      shading: { type: 'clear', fill: 'D9E2F3' },
      width: { size: Math.floor(9000 / colCount), type: WidthType.DXA },
    })),
  });

  // 创建数据行
  const dataRows = (rows || []).map(row =>
    new TableRow({
      children: row.map(cell =>
        new TableCell({
          children: [createParagraph(String(cell || ''), { size: 22, alignment: 'center' })],
          width: { size: Math.floor(9000 / colCount), type: WidthType.DXA },
        })
      ),
    })
  );

  const table = new Table({
    rows: [headerRow, ...dataRows],
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1 },
      bottom: { style: BorderStyle.SINGLE, size: 1 },
      left: { style: BorderStyle.SINGLE, size: 1 },
      right: { style: BorderStyle.SINGLE, size: 1 },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
      insideVertical: { style: BorderStyle.SINGLE, size: 1 },
    },
  });

  const elements = [];

  // 表名
  if (caption) {
    elements.push(createParagraph(caption, {
      size: 20, alignment: 'center', color: '555555', firstLineIndent: 0,
    }));
  }

  elements.push(table);
  elements.push(new Paragraph({ children: [], spacing: { after: 200 } }));
  return elements;
}

function createListSection(items, listType = 'bullet') {
  const results = [];
  (items || []).forEach((item, idx) => {
    results.push(new Paragraph({
      children: [createTextRun(String(item), { size: 24 })],
      bullet: listType === 'bullet' ? { level: 0 } : undefined,
      numbering: listType === 'number' ? { reference: 'main-numbering', level: 0 } : undefined,
      spacing: { after: 60, line: 340 },
      indent: { left: 480, hanging: 240 },
    }));
  });
  return results;
}

/**
 * 估算正文中文字数（不含空白和标点不影响统计）
 */
function estimateWordCount(chapters) {
  let count = 0;
  for (const ch of chapters) {
    count += ch.title.replace(/\s/g, '').length;
    if (Array.isArray(ch.content)) {
      for (const item of ch.content) {
        switch (item.type) {
          case 'paragraph':
          case 'heading':
            count += (item.text || '').replace(/\s/g, '').length;
            break;
          case 'list':
            (item.items || []).forEach(i => count += i.replace(/\s/g, '').length);
            break;
          case 'table':
            (item.headers || []).forEach(h => count += h.replace(/\s/g, '').length);
            (item.rows || []).forEach(row =>
              (row || []).forEach(cell => count += String(cell).replace(/\s/g, '').length)
            );
            count += (item.caption || '').replace(/\s/g, '').length;
            break;
          case 'image':
          case 'placeholder':
            count += (item.caption || '').replace(/\s/g, '').length;
            count += (item.title || '').replace(/\s/g, '').length;
            break;
        }
      }
    }
  }
  return count;
}

/**
 * 检测内容重复段落
 * 提取所有 paragraph 类型的 text，检查是否有重复
 */
function detectDuplicates(chapters) {
  const seen = {};
  const duplicates = [];
  for (const ch of chapters) {
    if (!Array.isArray(ch.content)) continue;
    for (const item of ch.content) {
      if (item.type === 'paragraph' && item.text) {
        // 取前40个字符作为key（足够长以区分相似但不同的段落，如项目概述 vs 章节引言）
        const key = item.text.replace(/\s/g, '').substring(0, 40);
        if (seen[key]) {
          duplicates.push({ first: seen[key], duplicate: item.text.substring(0, 60) + '...' });
        } else {
          seen[key] = item.text.substring(0, 60) + '...';
        }
      }
    }
  }
  return duplicates;
}

function buildContentFromChapter(chapter) {
  const elements = [];

  // 章节标题
  elements.push(createParagraph(chapter.title, {
    heading: chapter.level || 1,
    bold: true,
    size: chapter.level === 1 ? 32 : 28,
    alignment: 'left',
  }));

  // 章节内容
  if (Array.isArray(chapter.content)) {
    for (const item of chapter.content) {
      switch (item.type) {
        case 'paragraph':
          elements.push(createParagraph(item.text || '', item.format || {}));
          break;

        case 'heading':
          elements.push(createParagraph(item.text, {
            heading: item.level || 2,
            bold: true,
            size: item.level === 2 ? 28 : 26,
            alignment: 'left',
          }));
          break;

        case 'list': {
          const listItems = createListSection(item.items, item.listType || 'bullet');
          elements.push(...listItems);
          break;
        }

        case 'table': {
          const tableElements = createTableSection(item.headers, item.rows, item.caption, item.format);
          elements.push(...tableElements);
          break;
        }

        case 'pageBreak':
          elements.push(new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } }));
          break;

        case 'image': {
          if (item.base64) {
            const imgBuffer = Buffer.from(item.base64, 'base64');
            elements.push(new Paragraph({
              children: [new ImageRun({
                data: imgBuffer,
                transformation: {
                  width: item.width || 400,
                  height: item.height || 300,
                },
              })],
              alignment: AlignmentType.CENTER,
              spacing: { before: 200, after: 100 },
            }));
            // 图片标题
            if (item.caption) {
              elements.push(createParagraph(item.caption, {
                size: 20, alignment: 'center', color: '555555',
              }));
            }
          }
          break;
        }

        case 'placeholder': {
          // 图片占位符：生成带虚线边框的占位框 + 提示文字 + 图名
          const pw = item.width || 500;
          const ph = item.height || 350;
          const placeholderTitle = item.title || '【此处插入图片】';
          elements.push(new Paragraph({
            children: [],
            spacing: { before: 200, after: 0 },
          }));
          // 占位框 — 用单行表格模拟（虚线边框+浅灰底色）
          const placeholderTable = new Table({
            rows: [
              new TableRow({
                height: { value: Math.round(ph / 1.33), rule: 'atLeast' },
                children: [
                  new TableCell({
                    children: [createParagraph(placeholderTitle, {
                      size: 22, alignment: 'center', color: '888888',
                    })],
                    shading: { type: 'clear', fill: 'F5F5F5' },
                    width: { size: Math.round(pw / 6), type: WidthType.DXA },
                    verticalAlign: 'center',
                  }),
                ],
              }),
            ],
            width: { size: Math.round(pw / 6), type: WidthType.DXA },
            borders: {
              top: { style: BorderStyle.DASHED, size: 2, color: '999999' },
              bottom: { style: BorderStyle.DASHED, size: 2, color: '999999' },
              left: { style: BorderStyle.DASHED, size: 2, color: '999999' },
              right: { style: BorderStyle.DASHED, size: 2, color: '999999' },
            },
          });
          elements.push(placeholderTable);
          // 图片标题
          if (item.caption) {
            elements.push(createParagraph(item.caption, {
              size: 20, alignment: 'center', color: '555555',
              firstLineIndent: 0,
            }));
          }
          elements.push(new Paragraph({
            children: [],
            spacing: { after: 200 },
          }));
          break;
        }

        default:
          elements.push(createParagraph(String(item.text || JSON.stringify(item)), {}));
      }
    }
  }

  return elements;
}

// ========== 主函数 ==========

function buildDocument(data) {
  const {
    projectName = '工程名称',
    subtitle = '施工组织设计',
    companyName = '',
    date: dateStr = new Date().toLocaleDateString('zh-CN'),
    chapters = [],
  } = data;

  const children = [];

  // ---- 封面 ----
  children.push(...createCoverPage(projectName, subtitle, companyName, dateStr));

  // ---- 目录页 ----
  children.push(createParagraph('目  录', {
    bold: true, size: 32, alignment: 'center', color: COLOR_TITLE,
    font: { name: FONT_EN, eastAsia: FONT_TITLE },
  }));
  children.push(new Paragraph({ children: [], spacing: { after: 200 } }));

  for (const ch of chapters) {
    const indent = (ch.level || 1) > 1 ? '    ' : '';
    children.push(createParagraph(`${indent}${ch.title}`, {
      size: ch.level === 1 ? 26 : 24,
      bold: ch.level === 1,
      alignment: 'left',
      firstLineIndent: 0,
    }));
  }
  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ---- 正文 ----
  for (let i = 0; i < chapters.length; i++) {
    const chapter = chapters[i];
    // 非首章：在新页开始
    if (i > 0) {
      children.push(new Paragraph({ children: [new PageBreak()], spacing: { after: 0 } }));
    }
    const chapterElements = buildContentFromChapter(chapter);
    children.push(...chapterElements);

    // 章节间加空行
    children.push(new Paragraph({ children: [], spacing: { after: 100 } }));
  }

  // ---- 创建文档 ----
  const doc = new Document({
    title: projectName,
    description: `${projectName} 施工组织设计`,
    creator: 'Claude Code Skill',
    styles: {
      default: {
        document: {
          run: {
            font: FONT_EN,
            size: '24pt',
          },
          paragraph: {
            spacing: { after: 120, line: 360 },
          },
        },
        heading1: {
          run: {
            font: { name: FONT_EN, eastAsia: FONT_ZH },
            size: '32pt',
            bold: true,
            color: COLOR_TITLE,
          },
          paragraph: {
            spacing: { before: 240, after: 120 },
            outlineLevel: 0,
          },
        },
        heading2: {
          run: {
            font: { name: FONT_EN, eastAsia: FONT_ZH },
            size: '28pt',
            bold: true,
            color: '2E5090',
          },
          paragraph: {
            spacing: { before: 200, after: 100 },
            outlineLevel: 1,
          },
        },
        heading3: {
          run: {
            font: { name: FONT_EN, eastAsia: FONT_ZH },
            size: '26pt',
            bold: true,
            color: '2E5090',
          },
          paragraph: {
            spacing: { before: 160, after: 80 },
            outlineLevel: 2,
          },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 }, // A4
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          },
        },
        headers: {
          default: new Header({
            children: [new Paragraph({
              children: [createTextRun(subtitle || projectName || '施工组织设计', { size: 18, color: '888888' })],
              alignment: AlignmentType.RIGHT,
            })],
          }),
        },
        footers: {
          default: new Footer({
            children: [new Paragraph({
              children: [
                createTextRun('第 ', { size: 18, color: '888888' }),
                new TextRun({ children: [PageNumber.CURRENT], size: 18, color: '888888' }),
                createTextRun(' 页', { size: 18, color: '888888' }),
              ],
              alignment: AlignmentType.CENTER,
            })],
          }),
        },
        children,
      },
    ],
  });

  return doc;
}

// ========== CLI 入口 ==========

function main() {
  const args = process.argv.slice(2);
  let inputData = null;
  let outputPath = '';

  if (args.includes('--stdin')) {
    // 从 stdin 读取
    const chunks = [];
    process.stdin.on('data', chunk => chunks.push(chunk));
    process.stdin.on('end', () => {
      try {
        inputData = JSON.parse(Buffer.concat(chunks).toString('utf8'));
        outputPath = args[args.indexOf('--stdin') + 1] || '施工组织设计.docx';
        generateDoc(inputData, outputPath);
      } catch (e) {
        console.error('❌ 解析输入失败:', e.message);
        process.exit(1);
      }
    });
    return;
  }

  const inputIdx = args.indexOf('--input');
  const outputIdx = args.indexOf('--output');

  if (inputIdx === -1 && outputIdx === -1) {
    // 简单模式: 参数直接是 input.json output.docx
    if (args.length >= 1) {
      inputData = JSON.parse(fs.readFileSync(args[0], 'utf8'));
      outputPath = args[1] || '施工组织设计.docx';
      generateDoc(inputData, outputPath);
      return;
    }
    console.error('用法:');
    console.error('  node generate-docx.js --input <input.json> --output <output.docx>');
    console.error('  node generate-docx.js --stdin < output.json');
    console.error('  node generate-docx.js <input.json> [output.docx]');
    process.exit(1);
  }

  if (inputIdx !== -1) {
    const inputPath = args[inputIdx + 1];
    if (!inputPath) { console.error('请指定输入文件'); process.exit(1); }
    inputData = JSON.parse(fs.readFileSync(inputPath, 'utf8'));
  }

  if (outputIdx !== -1) {
    outputPath = args[outputIdx + 1] || '施工组织设计.docx';
  } else {
    outputPath = '施工组织设计.docx';
  }

  generateDoc(inputData, outputPath);
}

async function generateDoc(data, outputPath) {
  try {
    // === JSON 输入验证 ===
    const chapters = data.chapters || [];
    if (chapters.length === 0) {
      console.warn('⚠️  警告：chapters 数组为空，文档将只有封面和目录');
    }
    let validationErrors = [];
    chapters.forEach((ch, i) => {
      if (!ch.title) validationErrors.push(`第${i+1}章缺少 title 字段`);
      if (!Array.isArray(ch.content)) validationErrors.push(`第${i+1}章(${ch.title||'未知'})缺少 content 数组`);
      if (Array.isArray(ch.content)) {
        ch.content.forEach((item, j) => {
          if (!item.type) validationErrors.push(`第${i+1}章 content[${j}] 缺少 type 字段`);
          if (item.type === 'paragraph' && !item.text) validationErrors.push(`第${i+1}章 content[${j}] paragraph 缺少 text`);
          if (item.type === 'table' && (!item.headers || !item.rows)) validationErrors.push(`第${i+1}章 content[${j}] table 缺少 headers 或 rows`);
        });
      }
    });
    if (validationErrors.length > 0) {
      console.warn('⚠️  JSON 数据校验发现问题:');
      validationErrors.slice(0, 8).forEach(e => console.warn(`   - ${e}`));
      if (validationErrors.length > 8) console.warn(`   ... 还有 ${validationErrors.length - 8} 个问题`);
    }
    // === 校验结束 ===

    const totalChapters = chapters.length;
    const totalChars = estimateWordCount(chapters);
    console.log('📄 正在生成文档...');
    console.log(`   章节: ${totalChapters} 章, 字数: 约 ${totalChars} 字`);
    if (totalChars > 50000) {
      console.log('   大型文档处理中，请稍候...');
    }

    const genStart = Date.now();
    const doc = buildDocument(data);
    console.log(`   文档构建完成 (${(Date.now() - genStart) / 1000}s)`);

    const bufStart = Date.now();
    const buffer = await Packer.toBuffer(doc);

    // 字数估算
    const wordCount = estimateWordCount(data.chapters || []);
    const estPages = Math.ceil(wordCount / CHARS_PER_PAGE);

    // 重复检测
    const duplicates = detectDuplicates(data.chapters || []);
    if (duplicates.length > 0) {
      console.log('');
      console.log(`⚠️  发现 ${duplicates.length} 处可能重复的内容段:`);
      duplicates.slice(0, 5).forEach((d, i) => {
        console.log(`   ${i + 1}. "${d.duplicate}"`);
        console.log(`      首次出现: "${d.first}"`);
      });
      if (duplicates.length > 5) {
        console.log(`   ... 还有 ${duplicates.length - 5} 处`);
      }
      console.log('⚠️  建议：请检查并移除重复段落，重新生成');
      console.log('');
    }

    fs.writeFileSync(outputPath, buffer);
    const totalTime = ((Date.now() - genStart) / 1000).toFixed(1);
    const bufTime = ((Date.now() - bufStart) / 1000).toFixed(1);
    console.log(`   序列化完成 (${bufTime}s)`);
    console.log(`✅ 文档已生成: ${path.resolve(outputPath)}`);
    console.log(`   ${'='.repeat(30)}`);
    console.log(`   文件大小: ${(buffer.length / 1024).toFixed(1)} KB${buffer.length > 1048576 ? ' (' + (buffer.length / 1048576).toFixed(1) + ' MB)' : ''}`);
    console.log(`   正文字数: 约 ${wordCount} 字`);
    console.log(`   估算页数: 约 ${estPages} 页（含标题表格，不含封面目录）`);
    console.log(`   生成耗时: ${totalTime} 秒`);

    // 章节统计
    const sectionCount = (data.chapters || []).reduce((sum, ch) => {
      const subs = (ch.content || []).filter(c => c.type === 'heading' && c.level === 2).length;
      return sum + subs + 1;
    }, 0);
    const avgWordsPerSection = Math.round(wordCount / Math.max(1, sectionCount));
    // 统计表格和占位符数量
    const tableCount = chapters.reduce((sum, ch) => {
      return sum + (ch.content || []).filter(c => c.type === 'table').length;
    }, 0);
    const placeholderCount = chapters.reduce((sum, ch) => {
      return sum + (ch.content || []).filter(c => c.type === 'placeholder').length;
    }, 0);
    console.log(`   子章节数: ${sectionCount} 个`);
    console.log(`   平均每子章节: 约 ${avgWordsPerSection} 字`);
    console.log(`   表格数: ${tableCount} 个（每子章节 ${(tableCount / Math.max(1, sectionCount)).toFixed(1)} 个）`);
    console.log(`   图片占位符: ${placeholderCount} 个`);

    // 如果设置了目标字数或每节字数，给出对比
    if (data.targetWordCount) {
      const diff = wordCount - data.targetWordCount;
      const pct = Math.round((wordCount / data.targetWordCount) * 100);
      console.log(`   目标字数: ${data.targetWordCount} 字`);
      console.log(`   完成度: ${pct}%${diff >= 0 ? ' ✅' : ' ⚠️ 不足'}`);
    }
    if (data.targetWordsPerSection) {
      console.log(`   目标每子章节: ${data.targetWordsPerSection} 字`);
    }

    // === 智能建议 ===
    if (data.targetWordCount) {
      const pct = Math.round((wordCount / data.targetWordCount) * 100);
      const diff = wordCount - data.targetWordCount;
      if (pct < 85) {
        const neededChars = Math.abs(diff);
        const neededSubsections = Math.ceil(neededChars / 3500);
        console.log(`   💡 建议：当前完成度 ${pct}%，建议补充约 ${neededChars} 字（约 ${neededSubsections} 个子章节）`);
      }
    }
    if (avgWordsPerSection < 2000 && sectionCount > 0) {
      console.log(`   💡 建议：平均每子章节 ${avgWordsPerSection} 字偏低，可增加各子章节的细节深度`);
    }
    // === 建议结束 ===

    console.log(`   ${'='.repeat(30)}`);
  } catch (e) {
    console.error('❌ 生成文档失败:', e.message);
    console.error(e.stack);
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}

module.exports = { buildDocument, createParagraph, createTextRun, createListSection, createTableSection, createCoverPage, estimateWordCount, detectDuplicates };
