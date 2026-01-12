// 간단한 HTTP 서버
const http = require('http');
const fs = require('fs');
const path = require('path');
const CardDataAnalyzer = require('./card_analyzer.js');

const PORT = 3000;

const mimeTypes = {
    '.html': 'text/html',
    '.js': 'text/javascript',
    '.json': 'application/json',
    '.css': 'text/css'
};

const server = http.createServer((req, res) => {
    console.log(`${req.method} ${req.url}`);

    // API 엔드포인트
    if (req.url === '/api/analyze') {
        try {
            const analyzer = new CardDataAnalyzer();
            analyzer.load매입내역();
            analyzer.load포스결제내역();
            analyzer.load영수증내역();
            analyzer.load상품목록();
            analyzer.load매출상세();

            const 수수료통계 = analyzer.calculate실제수수료율();
            const 매칭결과 = analyzer.match매입포스();
            const 거래처정산 = analyzer.calculate거래처정산();

            // 거래처정산 - 품목별상세 포함 (아코디언용)
            const 거래처정산전체 = {};
            Object.entries(거래처정산).forEach(([거래처, data]) => {
                거래처정산전체[거래처] = data;
            });

            const result = {
                생성시각: new Date().toISOString(),
                요약: {
                    매입내역건수: analyzer.매입내역.length,
                    포스결제건수: analyzer.포스결제내역.length,
                    영수증품목수: analyzer.영수증내역.length,
                    매출상세건수: analyzer.매출상세.length,
                    거래처수: Object.keys(거래처정산).length
                },
                수수료통계: 수수료통계,
                매칭결과: {
                    매칭된항목: 매칭결과.매칭된항목.length,
                    포스누락: 매칭결과.포스누락,
                    매입누락: 매칭결과.매입누락,
                    상쇄된거래: 매칭결과.상쇄된거래 || []
                },
                매입내역: analyzer.매입내역,
                거래처정산: 거래처정산전체
            };

            res.writeHead(200, {
                'Content-Type': 'application/json',
                'Access-Control-Allow-Origin': '*'
            });
            res.end(JSON.stringify(result, null, 2));
        } catch (error) {
            console.error('API 오류:', error);
            res.writeHead(500, {
                'Content-Type': 'application/json',
                'Access-Control-Allow-Origin': '*'
            });
            res.end(JSON.stringify({ error: error.message, stack: error.stack }));
        }
        return;
    }

    // 정적 파일 서빙
    let filePath = '.' + req.url;
    if (filePath === './') {
        filePath = './dashboard.html';
    }

    const extname = path.extname(filePath);
    const contentType = mimeTypes[extname] || 'application/octet-stream';

    fs.readFile(filePath, (error, content) => {
        if (error) {
            if (error.code === 'ENOENT') {
                res.writeHead(404);
                res.end('404 Not Found');
            } else {
                res.writeHead(500);
                res.end('500 Internal Server Error: ' + error.code);
            }
        } else {
            res.writeHead(200, {
                'Content-Type': contentType,
                'Access-Control-Allow-Origin': '*'
            });
            res.end(content, 'utf-8');
        }
    });
});

server.listen(PORT, () => {
    console.log(`
========================================
  카드 결제 검증 대시보드 서버 시작!
========================================

  URL: http://localhost:${PORT}

  브라우저에서 위 주소로 접속하세요.
  종료하려면 Ctrl+C를 누르세요.

========================================
`);
});
