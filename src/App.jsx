import React, { useState, useEffect } from "react";
import pptxgen from "pptxgenjs";
import JSZip from "jszip";


// ==================================================
// 성경 약어 매핑 (66권 전체 지원)
// ==================================================
const BOOK_MAP = {
  // 구약성경
  "창": "창세기", "출": "출애굽기", "레": "레위기", "민": "민수기", "신": "신명기",
  "수": "여호수아", "삿": "사사기", "룻": "룻기",
  "삼상": "사무엘상", "삼하": "사무엘하",
  "왕상": "열왕기상", "왕하": "열왕기하",
  "대상": "역대상", "대하": "역대하",
  "스": "에스라", "느": "느헤미야", "에": "에스더",
  "욥": "욥기", "시": "시편", "잠": "잠언",
  "전": "전도서", "아": "아가",
  "사": "이사야", "렘": "예레미야", "애": "예레미야애가",
  "겔": "에스겔", "단": "다니엘",
  "호": "호세아", "욜": "요엘", "암": "아모스", "오": "오바디아", "욘": "요나", "미": "미가",
  "나": "나훔", "합": "하박국", "습": "스바냐", "학": "학개", "슥": "스가랴", "말": "말라기",

  // 신약성경
  "마": "마태복음", "막": "마가복음", "눅": "누가복음", "요": "요한복음",
  "행": "사도행전", "롬": "로마서", "고전": "고린도전서", "고후": "고린도후서",
  "갈": "갈라디아서", "엡": "에베소서", "빌": "빌립보서", "골": "골로새서",
  "살전": "데살로니가전서", "살후": "데살로니가후서", "딤전": "디모데전서", "딤후": "디모데후서",
  "딛": "디도서", "몬": "빌레몬서", "히": "히브리서", "야": "야고보서",
  "벧전": "베드로전서", "벧후": "베드로후서", "요일": "요한일서", "요이": "요한이서", "요삼": "요한삼서",
  "유": "유다서", "계": "요한계시록"
};

// ==================================================
// 샘플 성경 데이터 (즉시 체험용)
// ==================================================
const SAMPLE_BIBLE = {
  "창": {
    "1": {
      "1": "태초에 하나님이 천지를 창조하시니라",
      "2": "땅이 혼돈하고 공허하며 흑암이 깊음 위에 있고 하나님의 영은 수면 위에 운행하시니라",
      "3": "하나님이 이르시되 빛이 있으라 하시니 빛이 있었고"
    }
  },
  "시": {
    "23": {
      "1": "여호와는 나의 목자시니 내게 부족함이 없으리로다",
      "2": "그가 나를 푸른 풀밭에 누이시며 쉴 만한 물 가로 인도하시는도다",
      "3": "내 영혼을 소생시키시고 자기 이름을 위하여 의의 길로 인도하시는도다",
      "4": "내가 사망의 음침한 골짜기로 다닐지라도 해를 두려워하지 않을 것은 주께서 나와 함께 하심이라 주의 지팡이와 막대기가 나를 안위하시나이다",
      "5": "주께서 내 원수의 목전에서 내게 상을 차려 주시고 기름을 내 머리에 부으셨으니 내 잔이 넘치나이다",
      "6": "내 평생에 선하심과 인자하심이 반드시 나를 따르리니 내가 여호와의 집에 영원히 살리로다"
    }
  },
  "요": {
    "3": {
      "16": "하나님이 세상을 이처럼 사랑하사 독생자를 주셨으니 이는 그를 믿는 자마다 멸망하지 않고 영생을 얻게 하려 하심이라"
    }
  },
  "롬": {
    "8": {
      "28": "우리가 알거니와 하나님을 사랑하는 자 곧 그의 뜻대로 부르심을 입은 자들에게는 모든 것이 합력하여 선을 이루느니라"
    }
  }
};

// ==================================================
// 테마 설정 (프리미엄 슬라이드 스타일)
// ==================================================
const THEMES = [
  {
    id: "dark-purple",
    name: "미드나잇 퍼플 (기본)",
    className: "theme-dark",
    bgColor: "1e1a33",
    gradientColors: ["#1a1035", "#0f0c1a"],
    fontColor: "E2D9F7",
    accentColor: "F59E0B",
    fontSize: 28,
    fontFamily: "Noto Serif KR"
  },
  {
    id: "deep-blue",
    name: "딥 오션 블루",
    className: "theme-blue",
    bgColor: "0f2044",
    gradientColors: ["#0f2044", "#1a3a6e"],
    fontColor: "BFDBFE",
    accentColor: "38BDF8",
    fontSize: 28,
    fontFamily: "Noto Sans KR"
  },
  {
    id: "warm-gold",
    name: "클래식 워머 골드",
    className: "theme-gold",
    bgColor: "2d1b00",
    gradientColors: ["#2d1b00", "#4a2d00"],
    fontColor: "FDE68A",
    accentColor: "F59E0B",
    fontSize: 28,
    fontFamily: "Noto Serif KR"
  }
];

function App() {
  const [bible, setBible] = useState(SAMPLE_BIBLE);
  const [userInput, setUserInput] = useState("시 23:1-6, 요 3:16");
  const [filename, setFilename] = useState("");
  const [selectedTheme, setSelectedTheme] = useState(THEMES[0]);
  const [isGenerating, setIsGenerating] = useState(false);
  const [uploadedFileName, setUploadedFileName] = useState("");
  const [previewList, setPreviewList] = useState([]);
  const [toasts, setToasts] = useState([]);
  
  // PPT 템플릿 파일 상태
  const [templateFile, setTemplateFile] = useState(null);
  const [templateFileName, setTemplateFileName] = useState("");


  // 알림 토스트 추가
  const showToast = (message, type = "info") => {
    const id = Date.now();
    setToasts((prev) => [...prev, { id, message, type }]);
    setTimeout(() => {
      setToasts((prev) => prev.filter((t) => t.id !== id));
    }, 4000);
  };

  // 컴포넌트 마운트 시 public/bible_full.txt 자동 로드
  useEffect(() => {
    fetch("/bible_full.txt")
      .then((res) => {
        if (!res.ok) throw new Error("기본 성경 파일을 찾을 수 없습니다.");
        return res.text();
      })
      .then((text) => {
        const parsedBible = {};
        const lines = text.split("\n");
        let count = 0;

        lines.forEach((line) => {
          const match = line.trim().match(/^([가-힣]+)(\d+):(\d+)(.*)/);
          if (match) {
            const [_, book, chapter, verse, content] = match;
            if (!parsedBible[book]) parsedBible[book] = {};
            if (!parsedBible[book][chapter]) parsedBible[book][chapter] = {};
            parsedBible[book][chapter][verse] = content.trim();
            count++;
          }
        });

        if (count > 0) {
          setBible(parsedBible);
          setUploadedFileName("기본 전체 성경(창세기 등 포함)");
          showToast(`기본 성경 데이터 로드 완료! (${count.toLocaleString()}개 구절)`, "success");
        }
      })
      .catch((err) => {
        console.warn("기본 성경 자동 로딩 실패:", err);
      });
  }, []);


  // 입력 파싱 및 실시간 미리보기 리스트 생성
  useEffect(() => {
    if (!userInput.trim()) {
      setPreviewList([]);
      return;
    }
    const refs = parseInput(userInput);
    const resolved = [];

    refs.forEach(([book_short, chapter, verse]) => {
      const book_full = BOOK_MAP[book_short] || book_short;

      if (verse === "ALL") {
        if (bible[book_short] && bible[book_short][chapter]) {
          const sortedVerses = Object.keys(bible[book_short][chapter]).sort((a, b) => parseInt(a) - parseInt(b));
          sortedVerses.forEach((v) => {
            resolved.push({
              key: `${book_short} ${chapter}:${v}`,
              book: book_full,
              bookShort: book_short,
              chapter,
              verse: v,
              content: bible[book_short][chapter][v]
            });
          });
        }
      } else {
        if (bible[book_short] && bible[book_short][chapter] && bible[book_short][chapter][verse]) {
          resolved.push({
            key: `${book_short} ${chapter}:${verse}`,
            book: book_full,
            bookShort: book_short,
            chapter,
            verse,
            content: bible[book_short][chapter][verse]
          });
        } else {
          // 샘플/업로드 데이터에 없는 경우 placeholder 표시
          resolved.push({
            key: `${book_short} ${chapter}:${verse}`,
            book: book_full,
            bookShort: book_short,
            chapter,
            verse,
            content: `구절 정보를 찾을 수 없습니다. (전체 성경 텍스트 파일을 업로드해 주세요)`
          });
        }
      }
    });

    setPreviewList(resolved);
  }, [userInput, bible]);

  // 성경 텍스트 파일 파싱
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setUploadedFileName(file.name);
    const reader = new FileReader();

    reader.onload = (event) => {
      const text = event.target.result;
      const parsedBible = {};
      const lines = text.split("\n");
      let count = 0;

      lines.forEach((line) => {
        // [가-힣]+ 로 시작하고 숫자:숫자 내용 형식인 라인을 정규식으로 매치
        const match = line.trim().match(/^([가-힣]+)(\d+):(\d+)(.*)/);
        if (match) {
          const [_, book, chapter, verse, content] = match;
          if (!parsedBible[book]) parsedBible[book] = {};
          if (!parsedBible[book][chapter]) parsedBible[book][chapter] = {};
          parsedBible[book][chapter][verse] = content.trim();
          count++;
        }
      });

      if (count > 0) {
        setBible(parsedBible);
        showToast(`성경 로드 완료! 총 ${count.toLocaleString()}개 구절이 등록되었습니다.`, "success");
      } else {
        showToast("올바른 성경 텍스트 형식이 아닙니다. '창1:1 태초에...' 형식인지 확인해주세요.", "error");
        setUploadedFileName("");
      }
    };

    reader.readAsText(file, "utf-8");
  };

  // PPTX 템플릿 파일 로드
  const handleTemplateUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (!file.name.endsWith(".pptx")) {
      showToast("올바른 PowerPoint 파일(.pptx)을 업로드해주세요.", "error");
      return;
    }

    setTemplateFileName(file.name);
    const reader = new FileReader();

    reader.onload = (event) => {
      setTemplateFile(event.target.result); // ArrayBuffer
      showToast("PPTX 템플릿 양식이 성공적으로 등록되었습니다!", "success");
    };

    reader.readAsArrayBuffer(file);
  };


  // 입력 파서
  const parseInput = (user_input) => {
    const results = [];
    const parts = user_input.split(",");

    parts.forEach((part) => {
      const trimmed = part.trim();
      if (!trimmed) return;

      // 절 or 절범위 (예: 창 1:1-3, 시 23:1)
      const matchVerse = trimmed.match(/^([가-힣]+)\s*(\d+):([\d\-]+)/);
      if (matchVerse) {
        const [_, book, chapter, verse_part] = matchVerse;
        if (verse_part.includes("-")) {
          const [start, end] = verse_part.split("-");
          const startNum = parseInt(start);
          const endNum = parseInt(end);
          if (!isNaN(startNum) && !isNaN(endNum)) {
            for (let v = startNum; v <= endNum; v++) {
              results.push([book, chapter, String(v)]);
            }
          }
        } else {
          results.push([book, chapter, verse_part]);
        }
        return;
      }

      // 여러 장 범위 전체 (예: 창 1-3)
      const matchRange = trimmed.match(/^([가-힣]+)\s*(\d+)-(\d+)$/);
      if (matchRange) {
        const [_, book, start_ch, end_ch] = matchRange;
        const startNum = parseInt(start_ch);
        const endNum = parseInt(end_ch);
        if (!isNaN(startNum) && !isNaN(endNum)) {
          for (let ch = startNum; ch <= endNum; ch++) {
            results.push([book, String(ch), "ALL"]);
          }
        }
        return;
      }

      // 장 전체 (예: 시 23, 요 3)
      const matchChapter = trimmed.match(/^([가-힣]+)\s*(\d+)$/);
      if (matchChapter) {
        const [_, book, chapter] = matchChapter;
        results.push([book, chapter, "ALL"]);
      }
    });

    return results;
  };

  // PPT 생성 및 다운로드 (양식 업로드 시 JSZip 복사, 미업로드 시 PptxGenJS 테마 적용)
  const generatePPTX = async () => {
    if (previewList.length === 0) {
      showToast("생성할 성경 구절이 없습니다. 구절을 올바르게 입력해주세요.", "error");
      return;
    }

    setIsGenerating(true);
    showToast("파워포인트 슬라이드를 구성하고 있습니다...", "info");

    const today = new Date().toISOString().slice(0, 10).replace(/-/g, "");
    const outputName = (filename.trim() || today) + ".pptx";

    try {
      if (templateFile) {
        // ==================================================
        // 사용자가 PPTX 양식(템플릿)을 업로드한 경우 -> JSZip으로 슬라이드 복제 및 치환
        // ==================================================
        const zip = await JSZip.loadAsync(templateFile);
        
        // 메인 핵심 XML 파일들 로드
        let presentationXml = await zip.file("ppt/presentation.xml").async("string");
        let presentationRelsXml = await zip.file("ppt/_rels/presentation.xml.rels").async("string");
        let contentTypesXml = await zip.file("[Content_Types].xml").async("string");
        
        // 템플릿의 첫 번째 슬라이드 내용 로드
        const baseSlideXml = await zip.file("ppt/slides/slide1.xml").async("string");
        
        let baseSlideRelsXml = "";
        if (zip.file("ppt/slides/_rels/slide1.xml.rels")) {
          baseSlideRelsXml = await zip.file("ppt/slides/_rels/slide1.xml.rels").async("string");
        }

        let sldIdLstContent = "";
        let relsContent = "";
        let contentTypesOverrides = "";

        previewList.forEach((item, index) => {
          const slideIndex = index + 2; // 슬라이드 2번부터 새롭게 생성
          const rId = `rIdBibleSlide${slideIndex}`;
          const sldId = 256 + index + 1; // 고유 슬라이드 ID

          const chapterSuffix = item.bookShort === "시" ? "편" : "장";
          let slideXml = baseSlideXml;

          // XML 내부의 플레이스홀더 문자열 치환
          // 윈도우 한글 PPT 특성상 XML 태그가 텍스트 사이사이에 쪼개져 있을 수 있으므로 단순 텍스트 치환과 정규식 양용
          slideXml = slideXml
            .replace(/\(BOOK\)/g, item.book)
            .replace(/\(CHAPTER\)/g, `${item.chapter}${chapterSuffix}`)
            .replace(/\(VERSE\)/g, `${item.verse}절`)
            .replace(/\(CONTENT\)/g, item.content);

          // ZIP 파일 내에 개별 슬라이드 파일 생성
          zip.file(`ppt/slides/slide${slideIndex}.xml`, slideXml);
          if (baseSlideRelsXml) {
            zip.file(`ppt/slides/_rels/slide${slideIndex}.xml.rels`, baseSlideRelsXml);
          }

          // XML에 등록할 데이터 수집
          sldIdLstContent += `<p:sldId id="${sldId}" r:id="${rId}"/>`;
          relsContent += `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideIndex}.xml"/>`;
          contentTypesOverrides += `<Override PartName="/ppt/slides/slide${slideIndex}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
        });

        // presentation.xml에 신규 슬라이드 등록
        presentationXml = presentationXml.replace(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/, (match, inner) => {
          return `<p:sldIdLst>${inner}${sldIdLstContent}</p:sldIdLst>`;
        });

        // presentation.xml.rels에 신규 슬라이드 관계성 등록
        presentationRelsXml = presentationRelsXml.replace(/<\/Relationships>/, `${relsContent}</Relationships>`);

        // [Content_Types].xml에 신규 슬라이드 오버라이드 등록
        contentTypesXml = contentTypesXml.replace(/<\/Types>/, `${contentTypesOverrides}</Types>`);

        // 첫 번째 슬라이드(원본 템플릿 슬라이드)를 화면에서 제거 (파이썬의 sldIdLst.remove(list(sldIdLst)[0]) 로직 구현)
        if (previewList.length > 0) {
          presentationXml = presentationXml.replace(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/, (match, inner) => {
            // 첫 번째 <p:sldId .../> 태그 하나만 제거하여 템플릿 슬라이드 미출력 처리
            return `<p:sldIdLst>${inner.replace(/<p:sldId[^>]*?\/>/, "")}</p:sldIdLst>`;
          });
        }

        // 수정된 XML 파일들을 다시 ZIP 아카이브에 쓰기
        zip.file("ppt/presentation.xml", presentationXml);
        zip.file("ppt/_rels/presentation.xml.rels", presentationRelsXml);
        zip.file("[Content_Types].xml", contentTypesXml);

        // ZIP 파일을 바이너리 블롭으로 압축 및 다운로드 트리거
        const blob = await zip.generateAsync({ type: "blob" });
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = outputName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        showToast("업로드하신 PPT 양식으로 변환 다운로드가 완료되었습니다!", "success");
      } else {
        // ==================================================
        // 템플릿 미업로드 시 -> 기존의 아름다운 프리미엄 테마로 새 슬라이드 생성 (PptxGenJS)
        // ==================================================
        const pptx = new pptxgen();
        pptx.layout = "LAYOUT_16x9";

        previewList.forEach((item) => {
          const slide = pptx.addSlide();
          slide.background = { color: selectedTheme.bgColor };

          const chapterSuffix = item.bookShort === "시" ? "편" : "장";
          const refText = `${item.book} ${item.chapter}${chapterSuffix} ${item.verse}절`;

          slide.addText(refText, {
            x: 1.0,
            y: 0.8,
            w: "80%",
            h: 0.6,
            fontSize: 18,
            fontFace: selectedTheme.fontFamily,
            color: selectedTheme.accentColor,
            bold: true,
            align: "left"
          });

          slide.addShape(pptx.shapes.RECTANGLE, {
            x: 1.0,
            y: 1.4,
            w: 11.3,
            h: 0.02,
            fill: { color: selectedTheme.accentColor }
          });

          slide.addText(item.content, {
            x: 1.0,
            y: 2.0,
            w: 11.3,
            h: 4.2,
            fontSize: selectedTheme.fontSize,
            fontFace: selectedTheme.fontFamily,
            color: selectedTheme.fontColor,
            align: "center",
            valign: "middle",
            lineSpacing: 42,
            bold: selectedTheme.fontFamily.includes("Serif")
          });
        });

        await pptx.writeFile({ fileName: outputName });
        showToast("기본 테마 PPTX 파일 다운로드가 완료되었습니다!", "success");
      }
    } catch (error) {
      console.error(error);
      showToast("PPT 생성 또는 템플릿 변환 도중 오류가 발생했습니다.", "error");
    } finally {
      setIsGenerating(false);
    }
  };


  return (
    <div className="app">
      <header className="header">
        <div className="header-icon">📖</div>
        <h1>성경 PPT 생성기</h1>
        <p>
          깔끔하고 고품격 스타일의 성경 슬라이드를 클릭 한 번으로 쉽고 빠르게 제작하세요.
        </p>
      </header>

      <main className="main-container">
        {/* 1. PPT 양식(템플릿) 등록 카드 */}
        <div className="card">
          <h2 className="card-title">
            <span></span>1. PPT 템플릿 양식 등록
          </h2>
          
          {templateFileName ? (
            <div className="upload-success" style={{ marginBottom: "16px" }}>
              <span>✅</span>
              <span className="file-name">{templateFileName}</span>
              <span>양식 적용 완료</span>
              <button 
                onClick={() => { setTemplateFile(null); setTemplateFileName(""); }}
                style={{ marginLeft: "auto", background: "none", border: "none", color: "#fca5a5", cursor: "pointer", fontSize: "12px", fontWeight: "600" }}
              >
                취소
              </button>
            </div>
          ) : (
            <div className="upload-zone" style={{ marginBottom: "16px" }}>
              <span className="upload-icon">📊</span>
              <p className="upload-text">나만의 PPT 템플릿 파일(.pptx)을 업로드하세요</p>
              <p className="upload-hint">업로드하지 않을 시 아래 3번 테마 중 선택하신 프리미엄 기본 테마로 슬라이드가 생성됩니다.</p>
              <input type="file" accept=".pptx" onChange={handleTemplateUpload} />
            </div>
          )}

          <div className="syntax-grid" style={{ marginTop: "12px" }}>
            <div style={{ gridColumn: "1 / -1", fontSize: "13px", fontWeight: "600", color: "var(--accent-purple)", marginBottom: "4px" }}>
              💡 PPT 템플릿 제작 및 플레이스홀더 사용 규칙:
            </div>
            <div className="syntax-item">
              <div className="syntax-code">(BOOK)</div>
              <div className="syntax-desc">성경 책 전체 이름으로 치환됩니다. (예: 창세기, 요한복음)</div>
            </div>
            <div className="syntax-item">
              <div className="syntax-code">(CHAPTER)</div>
              <div className="syntax-desc">장 혹은 편 명칭으로 치환됩니다. (예: 1장, 23편)</div>
            </div>
            <div className="syntax-item">
              <div className="syntax-code">(VERSE)</div>
              <div className="syntax-desc">절 명칭으로 치환됩니다. (예: 1절, 16절)</div>
            </div>
            <div className="syntax-item">
              <div className="syntax-code">(CONTENT)</div>
              <div className="syntax-desc">선택된 성경 구절의 실제 본문 내용으로 치환됩니다.</div>
            </div>
            <div style={{ gridColumn: "1 / -1", fontSize: "12px", color: "var(--text-muted)", marginTop: "6px", lineHeight: "1.4" }}>
              ※ 템플릿의 <strong>첫 번째 슬라이드</strong>가 원본 양식으로 활용되며, 파싱된 모든 구절 수만큼 이 양식을 바탕으로 슬라이드가 자동 복제되어 합쳐집니다.
            </div>
          </div>
        </div>



        {/* 2. 구절 및 파일명 설정 카드 */}
        <div className="card">
          <h2 className="card-title">
            <span></span>2. 구절 입력 & 설정
          </h2>

          <div className="form-group">
            <label htmlFor="ref-input">
              구절 입력
              <span className="label-badge">쉼표(,)로 다중 입력 가능</span>
            </label>
            <input
              id="ref-input"
              type="text"
              placeholder="예: 시 23:1-6, 요 3:16"
              value={userInput}
              onChange={(e) => setUserInput(e.target.value)}
            />
            
            <div className="syntax-grid">
              <div className="syntax-item">
                <div className="syntax-code">시 23:1-6</div>
                <div className="syntax-desc">시편 23편 1절부터 6절까지</div>
              </div>
              <div className="syntax-item">
                <div className="syntax-code">요 3:16</div>
                <div className="syntax-desc">요한복음 3장 16절 단일 구절</div>
              </div>
              <div className="syntax-item">
                <div className="syntax-code">창 1</div>
                <div className="syntax-desc">창세기 1장 전체 구절</div>
              </div>
              <div className="syntax-item">
                <div className="syntax-code">창 1-3</div>
                <div className="syntax-desc">창세기 1장부터 3장 전체</div>
              </div>
            </div>
          </div>

          <div className="form-group">
            <label htmlFor="filename-input">파일명</label>
            <input
              id="filename-input"
              type="text"
              placeholder="비워둘 시 오늘 날짜로 자동 지정됩니다"
              value={filename}
              onChange={(e) => setFilename(e.target.value)}
            />
          </div>
        </div>

        {/* 3. 디자인 테마 선택 카드 */}
        <div className="card">
          <h2 className="card-title">
            <span></span>3. 슬라이드 테마 선택
          </h2>
          <div className="theme-grid">
            {THEMES.map((t) => (
              <label key={t.id} className="theme-option">
                <input
                  type="radio"
                  name="theme"
                  checked={selectedTheme.id === t.id}
                  onChange={() => setSelectedTheme(t)}
                />
                <div className={`theme-preview ${t.className}`}>
                  <div className="theme-title-preview" style={{ color: t.accentColor }}>
                    창세기 1장 1절
                  </div>
                  <div className="theme-text-preview" style={{ color: t.fontColor }}>
                    태초에 하나님이 천지를 창조하시니라.
                  </div>
                </div>
                <div className="theme-label">{t.name}</div>
              </label>
            ))}
          </div>
        </div>

        {/* 4. 실시간 프리뷰 카드 */}
        <div className="card">
          <h2 className="card-title">
            <span></span>실시간 구절 매칭 및 미리보기
          </h2>
          <div className="preview-list">
            {previewList.length > 0 ? (
              previewList.map((item, idx) => (
                <div key={idx} className="preview-item">
                  <span className="preview-badge">
                    {item.book} {item.chapter}:{item.verse}
                  </span>
                  <span className="preview-verse">{item.content}</span>
                </div>
              ))
            ) : (
              <div className="preview-empty">
                입력된 구절이 없거나 형식이 맞지 않습니다.
              </div>
            )}
          </div>
          {previewList.length > 0 && (
            <div className="preview-count">
              총 <strong>{previewList.length}</strong>개의 슬라이드가 생성될 예정입니다.
            </div>
          )}
        </div>

        {/* PPT 생성 버튼 */}
        <button
          className="btn-generate"
          onClick={generatePPTX}
          disabled={isGenerating || previewList.length === 0}
        >
          {isGenerating ? (
            <>
              <div className="spinner"></div>
              PPT 생성 중...
            </>
          ) : (
            <>
              <span>📊</span>
              PPT 파일 생성 및 다운로드
            </>
          )}
        </button>
      </main>

      <footer className="footer">
        © {new Date().getFullYear()} 성경 PPT 생성기. All rights reserved.
      </footer>

      {/* 토스트 컨테이너 */}
      <div className="toast-container">
        {toasts.map((t) => (
          <div key={t.id} className={`toast ${t.type}`}>
            {t.type === "success" && "✨ "}
            {t.type === "error" && "⚠️ "}
            {t.message}
          </div>
        ))}
      </div>
    </div>
  );
}

export default App;
