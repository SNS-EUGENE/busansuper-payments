// 카드 결제 데이터 분석 및 검증 프로그램
const XLSX = require('xlsx');
const fs = require('fs');

// 카드사별 수수료율 정보 (부산슈퍼 결제사 수수료율.xlsx 기반)
// 수수료율 단위: % (예: 2.3 = 2.3%)
const CARD_FEE_INFO = {
    'KB카드': {
        신용카드: 2.3,
        체크카드: 1.5,
        해외VISA: 4.0,
        AMEX: 2.3,
        기타: '카뱅체크 1.35%',
        정산일: 'T+2'
    },
    '국민카드': {  // KB카드와 동일
        신용카드: 2.3,
        체크카드: 1.5,
        해외VISA: 4.0,
        AMEX: 2.3,
        기타: '카뱅체크 1.35%',
        정산일: 'T+2'
    },
    '신한카드': {
        신용카드: 2.3,
        체크카드: 1.5,
        해외VISA: 3.5,
        정산일: 'T+2'
    },
    'BC카드': {
        신용카드: 2.3,
        체크카드: 1.35,
        해외VISA: 4.5,
        기타: 'NAPAS/SacomPay/VietPay/MIR/GPN/EVONET 4.50%',
        정산일: 'T+2'
    },
    '롯데카드': {
        신용카드: 2.3,
        체크카드: 1.5,
        해외VISA: 3.4,
        정산일: 'T+2'
    },
    '삼성카드': {
        신용카드: 2.3,
        체크카드: 1.55,
        해외VISA: 3.5,
        정산일: 'T+2'
    },
    '현대카드': {
        신용카드: 2.3,
        체크카드: 1.55,
        해외VISA: 4.0,
        정산일: 'T+2'
    },
    '농협카드': {
        신용카드: 2.3,
        체크카드: 1.45,
        해외VISA: 4.0,
        정산일: 'T+2'
    },
    '하나카드': {
        신용카드: 2.3,
        체크카드: 1.51,
        해외VISA: 3.6,
        기타: 'CUP 3.60%',
        정산일: 'T+2'
    },
    '우리카드': {
        신용카드: 2.3,
        체크카드: 1.52,
        정산일: 'T+2'
    },
    '카카오페이': {
        머니: 1.8,
        정산일: 'T+2'
    },
    '알리페이': {
        신용카드: 1.8,
        정산일: 'D+5'
    },
    '알리페이플러스': {
        신용카드: 1.8,
        정산일: 'D+5'
    },
    '위챗페이': {
        신용카드: 1.8,
        정산일: 'D+5'
    },
    '라인페이': {
        신용카드: 3.5,
        정산일: 'T+2'
    }
};

class CardDataAnalyzer {
    constructor() {
        this.매입내역 = [];
        this.포스결제내역 = [];
        this.영수증내역 = [];
        this.매출상세 = [];        // 영수증별매출상세현황
        this.상품목록 = [];        // 상품목록 (거래처 매핑용)
        this.상품Map = new Map();  // 바코드/상품코드 → 거래처 매핑
        this.검증결과 = {
            수수료율검증: [],
            매칭검증: [],
            누락내역: [],
            중복내역: []
        };
    }

    // 카드사명 정규화 (BC카드/비씨카드 등 통합)
    normalize카드사(카드사명) {
        if (!카드사명) return '';
        const normalized = String(카드사명).trim();

        // BC카드 / 비씨카드 통합
        if (normalized.includes('BC') || normalized.includes('비씨')) {
            return 'BC카드';
        }

        // 국민카드 → KB카드 통합
        if (normalized.includes('국민')) {
            return 'KB카드';
        }

        return normalized;
    }

    // 카드 유형 판별 (신용카드/체크카드)
    // 카드번호 BIN (앞 4-6자리)으로 판별
    detect카드유형(카드번호, 카드사) {
        if (!카드번호) return '신용카드';  // 기본값

        const cardNum = String(카드번호).replace(/[-*\s]/g, '');
        const bin = cardNum.substring(0, 4);

        // 체크카드 BIN 패턴 (일반적인 패턴)
        // 실제로는 각 카드사별 BIN 테이블이 필요하지만,
        // 간단히 일반적인 패턴으로 판별
        const 체크카드BIN = {
            'KB카드': ['5365', '9490', '4265'],  // KB체크 패턴
            '신한카드': ['5412', '9410'],
            'BC카드': ['4561', '9420'],
            '하나카드': ['4569'],
            '삼성카드': ['9440'],
            '현대카드': ['9450'],
            '우리카드': ['4023', '9430'],
            '농협카드': ['9460'],
            '롯데카드': ['9470']
        };

        // 실제로는 매입 파일의 '상품구분' 또는 '카드구분' 컬럼이 있으면 사용
        // 여기서는 간단히 기본값 반환
        return '신용카드';
    }

    // 수수료율 조회
    get수수료율(카드사, 카드유형 = '신용카드') {
        const 정규화된카드사 = this.normalize카드사(카드사);
        const info = CARD_FEE_INFO[정규화된카드사];

        if (!info) {
            return { 수수료율: 2.3, 정산일: 'T+2', 카드유형: '신용카드' };  // 기본값
        }

        let 수수료율 = info.신용카드 || 2.3;

        if (카드유형 === '체크카드' && info.체크카드) {
            수수료율 = info.체크카드;
        } else if (카드유형 === '머니' && info.머니) {
            수수료율 = info.머니;
        } else if (카드유형 === '해외' && info.해외VISA) {
            수수료율 = info.해외VISA;
        }

        return {
            수수료율: 수수료율,
            정산일: info.정산일 || 'T+2',
            카드유형: 카드유형,
            기타: info.기타 || ''
        };
    }

    // 수수료 계산
    calculate수수료(금액, 카드사, 카드유형 = '신용카드') {
        const feeInfo = this.get수수료율(카드사, 카드유형);
        const 수수료 = Math.round(금액 * feeInfo.수수료율 / 100);

        return {
            원금: 금액,
            수수료: 수수료,
            정산금액: 금액 - 수수료,
            수수료율: feeInfo.수수료율,
            카드유형: feeInfo.카드유형,
            정산일: feeInfo.정산일
        };
    }

    // 1. 기간별 매입내역 파싱
    load매입내역(filename = '여신_매입내역.xlsx') {
        const workbook = XLSX.readFile(filename);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        // 헤더 찾기 (No., 거래일자로 시작하는 행)
        let headerIdx = -1;
        for (let i = 0; i < rawData.length; i++) {
            if (rawData[i][0] === 'No.' || rawData[i][0] === 'NO.') {
                headerIdx = i;
                break;
            }
        }

        if (headerIdx === -1) {
            throw new Error('매입내역 헤더를 찾을 수 없습니다.');
        }

        const headers = rawData[headerIdx];
        const dataRows = rawData.slice(headerIdx + 1);

        this.매입내역 = dataRows
            .filter(row => row[0] && row[0] !== '' && !isNaN(row[0])) // No. 가 숫자인 행만
            .map(row => {
                const obj = {};
                headers.forEach((header, idx) => {
                    obj[header] = row[idx] || '';
                });
                // 카드사명 정규화
                if (obj['카드사']) {
                    obj['카드사'] = this.normalize카드사(obj['카드사']);
                }
                return obj;
            });

        console.log(`✓ 매입내역 ${this.매입내역.length}건 로드 완료`);
        return this.매입내역;
    }

    // 2. 포스 카드결제내역 파싱
    load포스결제내역(filename = '포스_카드결제내역.xlsx') {
        const workbook = XLSX.readFile(filename);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        // 헤더 찾기
        let headerIdx = -1;
        for (let i = 0; i < rawData.length; i++) {
            if (rawData[i][0] === 'No.' || rawData[i][0] === 'NO.') {
                headerIdx = i;
                break;
            }
        }

        if (headerIdx === -1) {
            throw new Error('포스결제내역 헤더를 찾을 수 없습니다.');
        }

        // 데이터 행 시작 (헤더가 2행 구조이므로 +2)
        const dataRows = rawData.slice(headerIdx + 2);

        // 실제 데이터 분석 결과 기반 하드코딩된 컬럼 매핑
        // Row 118 분석: [0]=113, [1]=45916(일자), [2]=01(포스), [3]=0004(영수증),
        //               [4]=승인(구분), [6]=국민카드, [7]=카드번호, [8]=요청금액,
        //               [10]=181(승인번호!), [11]=일시불, [14]=45916(승인일자),
        //               [15]=17:03:26(승인시각), [17]=2000(승인금액)
        this.포스결제내역 = dataRows
            .filter(row => row[0] && row[0] !== '' && !isNaN(row[0])) // No. 가 숫자인 행만
            .map(row => {
                return {
                    'No.': row[0],
                    '영업일자': row[1],
                    '포스번호': row[2],
                    '영수증번호': row[3],
                    '승인_구분': row[4],         // "승인" or "취소"
                    '승인_처리': row[5],         // "단말기승인" etc.
                    '매입사': this.normalize카드사(row[6] || ''),
                    '카드번호': row[7],
                    '승인요청금액': row[8],
                    '부가세': row[9],
                    '승인번호': row[10],         // ★ 중요: 실제 승인번호는 [10]에 있음!
                    '할부': row[11],
                    '할부개월': row[12],
                    '유효기간': row[13],
                    '승인일자': row[14],
                    '승인시각': row[15],
                    '승인금액': row[17]          // 최종 승인금액
                };
            });

        console.log(`✓ 포스결제내역 ${this.포스결제내역.length}건 로드 완료`);
        return this.포스결제내역;
    }

    // 3. 영수증별 매출 파싱
    load영수증내역(filename = '250910-251228 영수증별매출상세현황.xlsx') {
        const workbook = XLSX.readFile(filename);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        // 헤더 찾기
        let headerIdx = -1;
        for (let i = 0; i < rawData.length; i++) {
            if (rawData[i][0] === '일자') {
                headerIdx = i;
                break;
            }
        }

        if (headerIdx === -1) {
            throw new Error('영수증내역 헤더를 찾을 수 없습니다.');
        }

        const headers = rawData[headerIdx];
        const dataRows = rawData.slice(headerIdx + 1);

        this.영수증내역 = dataRows
            .filter(row => row[0] && row[0] !== '' && row[0].includes('-')) // 날짜 형식인 행만
            .map(row => {
                const obj = {};
                headers.forEach((header, idx) => {
                    obj[header] = row[idx] || '';
                });
                return obj;
            });

        console.log(`✓ 영수증내역 ${this.영수증내역.length}건 로드 완료`);
        return this.영수증내역;
    }

    // 3-1. 상품목록 파싱 (거래처 매핑용)
    load상품목록(filename = '부산슈퍼 상품목록.xlsx') {
        const workbook = XLSX.readFile(filename);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        // 헤더 찾기 (2번째 행)
        const headers = rawData[1];
        const dataRows = rawData.slice(2);

        this.상품목록 = dataRows
            .filter(row => row[0] && row[0] !== '')
            .map(row => {
                const obj = {};
                headers.forEach((header, idx) => {
                    obj[header] = row[idx] || '';
                });
                return obj;
            });

        // 바코드 → 거래처 매핑 구축
        this.상품Map.clear();
        this.상품목록.forEach(item => {
            const 바코드 = String(item['바코드'] || '').trim();
            const 상품코드 = String(item['상품코드'] || '').trim();
            const 거래처 = item['거래처'] || '미지정';
            const 판매단가 = parseFloat(item['판매단가']) || 0;

            if (바코드) {
                this.상품Map.set(바코드, { 거래처, 판매단가, 상품명: item['상품명'] });
            }
            if (상품코드) {
                this.상품Map.set(상품코드, { 거래처, 판매단가, 상품명: item['상품명'] });
            }
        });

        console.log(`✓ 상품목록 ${this.상품목록.length}건 로드 완료 (거래처 매핑: ${this.상품Map.size}개)`);
        return this.상품목록;
    }

    // 3-2. 영수증별매출상세현황 파싱
    load매출상세(filename = '250910-251228 영수증별매출상세현황.xlsx') {
        const workbook = XLSX.readFile(filename);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        // 헤더 찾기 (3번째 행)
        const headers = rawData[2];
        const dataRows = rawData.slice(3);

        // 1차 pass: 영수증별 할인합계 Map 구축
        const 영수증별할인합계 = new Map();
        let temp일자 = '';
        let temp영수증번호 = '';

        dataRows.forEach(row => {
            if (row[0] === '합계' || row[0] === '총합계' || !row[7]) return;

            const obj = {};
            headers.forEach((header, idx) => {
                obj[header] = row[idx] !== '' ? row[idx] : '';
            });

            if (obj['일자']) temp일자 = obj['일자'];
            if (obj['영수증번호']) temp영수증번호 = obj['영수증번호'];

            const 영수증키 = `${temp일자}_${temp영수증번호}`;
            const 할인액 = parseInt(obj['할인액']) || 0;

            if (!영수증별할인합계.has(영수증키)) {
                영수증별할인합계.set(영수증키, 0);
            }
            영수증별할인합계.set(영수증키, 영수증별할인합계.get(영수증키) + 할인액);
        });

        // 영수증별 할인합계 저장 (analyze할인유형에서 참조)
        this.영수증별할인합계 = 영수증별할인합계;
        console.log(`✓ 영수증별 할인합계 계산 완료 (${영수증별할인합계.size}건)`);

        // 2차 pass: 본격 파싱 + 할인유형분석
        let current일자 = '';
        let current영수증번호 = '';

        this.매출상세 = [];
        dataRows.forEach(row => {
            // 합계 행 또는 빈 행 스킵
            if (row[0] === '합계' || row[0] === '총합계' || !row[7]) return;

            const obj = {};
            headers.forEach((header, idx) => {
                obj[header] = row[idx] !== '' ? row[idx] : '';
            });

            // 빈 값은 이전 영수증 정보 유지
            if (obj['일자']) current일자 = obj['일자'];
            if (obj['영수증번호']) current영수증번호 = obj['영수증번호'];

            obj['일자'] = current일자;
            obj['영수증번호'] = current영수증번호;

            // 거래처 매핑
            const 바코드 = String(obj['바코드'] || '').trim();
            const 상품코드 = String(obj['상품코드'] || '').trim();
            const 상품정보 = this.상품Map.get(바코드) || this.상품Map.get(상품코드);

            obj['거래처'] = 상품정보 ? 상품정보.거래처 : '미지정';
            obj['매입가'] = 상품정보 ? 상품정보.판매단가 : parseInt(obj['총매출액']) || 0;

            // 할인 유형 분석 (영수증별 할인합계 전달)
            const 영수증키 = `${obj['일자']}_${obj['영수증번호']}`;
            const 영수증할인합계 = 영수증별할인합계.get(영수증키) || 0;
            obj['할인유형분석'] = this.analyze할인유형(obj, 영수증할인합계);

            this.매출상세.push(obj);
        });

        console.log(`✓ 매출상세 ${this.매출상세.length}건 로드 완료`);
        return this.매출상세;
    }

    // 할인 유형 분석 (쿠폰/서비스/일반 + 부담주체)
    // 정산 규칙:
    // - 쿠폰할인: 판매자(우리) 부담 → 총매출액 기준 정산
    // - 서비스할인 (2+1 무료분): 거래처 부담 → 정산 0원
    // - 일반할인 - 비율할인(15%, 20%): 거래처 부담 → 실매출액 기준 정산
    // - 일반할인 - 1+1 행사업체 50% 할인: 거래처 부담 → 실매출액 기준 정산
    // - 일반할인 - 금액권할인(1000,2000,3000,5000,10000원): 판매자 부담 → 총매출액 기준
    // - 일반할인 - 기타: 판매자(우리) 부담 → 총매출액 기준 정산
    analyze할인유형(row, 영수증할인합계 = 0) {
        const 할인구분 = row['할인구분'] || '';
        const 총매출액 = parseInt(row['총매출액']) || 0;
        const 할인액 = parseInt(row['할인액']) || 0;
        const 실매출액 = parseInt(row['실매출액']) || 0;
        const 거래처 = row['거래처'] || '';

        // 1+1 행사 업체 목록 (50% 할인 시 거래처 부담)
        const 행사업체목록 = ['해승J&T', '부산맥주', '모모스커피', '까사부사노', '카페385'];

        // 할인율 계산
        const 할인율 = 총매출액 > 0 ? (할인액 / 총매출액 * 100) : 0;

        let 할인유형 = '할인없음';
        let 부담주체 = '없음';
        let 정산금액 = 실매출액;  // 기본값

        if (할인액 === 0) {
            return { 할인유형, 부담주체, 할인율: 0, 정산금액: 총매출액 };
        }

        if (할인구분 === '쿠폰할인') {
            // 쿠폰: 판매자(우리) 부담
            할인유형 = '쿠폰할인';
            부담주체 = '판매자';
            정산금액 = 총매출액;  // 원가 기준 정산 (할인액은 우리가 부담)
        } else if (할인구분 === '서비스할인') {
            // 서비스(100% 무료): 거래처 부담 (2+1, 1+1)
            할인유형 = '서비스할인';
            부담주체 = '거래처';
            정산금액 = 0;  // 무료 제공분은 거래처 정산 0
        } else if (할인구분 === '일반할인') {
            // 일반할인: 영수증 단위 금액권 vs 비율 할인 구분

            // 1단계: 영수증 할인합계가 금액권 금액인지 확인
            // 금액권: 1000, 2000, 3000, 5000, 10000원 (판매자 부담)
            const 금액권금액 = [1000, 2000, 3000, 5000, 10000];
            const is금액권할인 = 금액권금액.includes(영수증할인합계);

            if (is금액권할인) {
                // 금액권 할인: 판매자(우리) 부담
                할인유형 = `금액권할인(${영수증할인합계.toLocaleString()}원)`;
                부담주체 = '판매자';
                정산금액 = 총매출액;  // 우리가 할인 부담 → 총매출액 기준 정산
            } else {
                // 2단계: 비율 할인 여부 확인
                const is15할인 = Math.abs(할인율 - 15) < 2;
                const is20할인 = Math.abs(할인율 - 20) < 2;
                const is50할인 = Math.abs(할인율 - 50) < 2;
                const is행사업체 = 행사업체목록.includes(거래처);

                if (is15할인 || is20할인) {
                    // 15%, 20% 비율 할인: 거래처 부담
                    부담주체 = '거래처';
                    할인유형 = is20할인 ? '일반할인(20%)' : '일반할인(15%)';
                    정산금액 = 실매출액;  // 거래처가 할인 부담 → 실매출액 기준
                } else if (is50할인 && is행사업체) {
                    // 1+1 행사업체의 50% 할인: 거래처 부담 (1+1 행사로 간주)
                    할인유형 = '일반할인(50%-1+1)';
                    부담주체 = '거래처';
                    정산금액 = 실매출액;  // 거래처가 할인 부담 → 실매출액 기준
                } else {
                    // 기타 할인: 판매자(우리) 부담
                    할인유형 = `기타할인(${할인율.toFixed(0)}%/${할인액.toLocaleString()}원)`;
                    부담주체 = '판매자';
                    정산금액 = 총매출액;  // 우리가 할인 부담 → 총매출액 기준 정산
                }
            }
        }

        return {
            할인유형,
            부담주체,
            할인율: Math.round(할인율),
            정산금액,
            영수증할인합계  // 디버깅용
        };
    }

    // 엑셀 날짜 변환 (클래스 메서드)
    excelDateToYYYYMMDD(serial) {
        if (!serial) return '';
        if (typeof serial === 'string') {
            // YYYY-MM-DD 형식이면 그대로 반환
            if (serial.includes('-')) return serial;
            // YYYYMMDD 형식이면 YYYY-MM-DD로 변환
            if (/^\d{8}$/.test(serial)) {
                return `${serial.substring(0, 4)}-${serial.substring(4, 6)}-${serial.substring(6, 8)}`;
            }
            return serial;
        }
        if (typeof serial === 'number') {
            const epoch = new Date(1899, 11, 30);
            const date = new Date(epoch.getTime() + serial * 86400000);
            const year = date.getFullYear();
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const day = String(date.getDate()).padStart(2, '0');
            return `${year}-${month}-${day}`;
        }
        return String(serial);
    }

    // 카드 영수증 Set 구축 (날짜_영수증번호)
    build카드영수증Set() {
        const excelDateToYYYYMMDD = (serial) => {
            if (!serial) return '';
            if (typeof serial === 'string') return serial;
            if (typeof serial === 'number') {
                const epoch = new Date(1899, 11, 30);
                const date = new Date(epoch.getTime() + serial * 86400000);
                const year = date.getFullYear();
                const month = String(date.getMonth() + 1).padStart(2, '0');
                const day = String(date.getDate()).padStart(2, '0');
                return `${year}-${month}-${day}`;
            }
            return String(serial);
        };

        const 카드영수증Set = new Set();
        this.포스결제내역.forEach(row => {
            const 날짜 = excelDateToYYYYMMDD(row['영업일자']);
            const 영수증번호 = String(row['영수증번호'] || '').trim();
            if (날짜 && 영수증번호) {
                카드영수증Set.add(`${날짜}_${영수증번호}`);
            }
        });

        return 카드영수증Set;
    }

    // 거래처별 정산 계산 (카드/비카드 구분)
    calculate거래처정산(카드수수료율 = 2.3) {
        console.log('\n=== 거래처별 정산 계산 ===');

        // 카드 영수증 Set 구축
        const 카드영수증Set = this.build카드영수증Set();
        console.log(`카드 결제 영수증: ${카드영수증Set.size}건`);

        const 거래처별정산 = {};
        let 카드품목수 = 0;
        let 비카드품목수 = 0;

        this.매출상세.forEach(row => {
            const 거래처 = row['거래처'] || '미지정';
            const 분석 = row['할인유형분석'];
            const 정산금액기준 = 분석.정산금액;

            // 카드/비카드 판별
            const 영수증키 = `${row['일자']}_${row['영수증번호']}`;
            const 카드결제 = 카드영수증Set.has(영수증키);

            // 카드수수료: 카드결제만 적용
            const 카드수수료 = 카드결제 ? Math.round(정산금액기준 * 카드수수료율 / 100) : 0;
            const 최종정산액 = 정산금액기준 - 카드수수료;

            if (카드결제) 카드품목수++;
            else 비카드품목수++;

            if (!거래처별정산[거래처]) {
                거래처별정산[거래처] = {
                    거래처,
                    총매출: 0,
                    총할인: 0,
                    실매출: 0,
                    카드매출: 0,
                    비카드매출: 0,
                    판매자부담할인: 0,  // 쿠폰
                    거래처부담할인: 0,  // 일반/서비스 할인
                    카드수수료합계: 0,
                    정산예정액: 0,
                    품목수: 0,
                    카드품목수: 0,
                    비카드품목수: 0,
                    품목별상세: []
                };
            }

            const stat = 거래처별정산[거래처];
            stat.총매출 += parseInt(row['총매출액']) || 0;
            stat.총할인 += parseInt(row['할인액']) || 0;
            stat.실매출 += parseInt(row['실매출액']) || 0;
            stat.카드수수료합계 += 카드수수료;
            stat.정산예정액 += 최종정산액;
            stat.품목수++;

            if (카드결제) {
                stat.카드매출 += parseInt(row['실매출액']) || 0;
                stat.카드품목수++;
            } else {
                stat.비카드매출 += parseInt(row['실매출액']) || 0;
                stat.비카드품목수++;
            }

            if (분석.부담주체 === '판매자') {
                stat.판매자부담할인 += parseInt(row['할인액']) || 0;
            } else if (분석.부담주체 === '거래처') {
                stat.거래처부담할인 += parseInt(row['할인액']) || 0;
            }

            // 상품 정보 조회
            const 상품코드 = row['상품코드'] || '';
            const 바코드 = row['바코드'] || '';
            const 상품정보 = this.상품Map.get(바코드) || this.상품Map.get(상품코드) || {};
            const 판매단가 = parseInt(상품정보.판매단가) || 0;
            const 총매출 = parseInt(row['총매출액']) || 0;
            const 할인액 = parseInt(row['할인액']) || 0;
            const 할인율 = 총매출 > 0 ? Math.round((할인액 / 총매출) * 100) : 0;

            // 카드 결제인 경우 카드사 정보 조회
            let 카드사명 = '';
            if (카드결제) {
                // 포스결제내역에서 해당 영수증의 카드사 찾기
                const 포스결제 = this.포스결제내역.find(p => {
                    const 포스날짜 = this.excelDateToYYYYMMDD(p['영업일자']);
                    return 포스날짜 === row['일자'] && String(p['영수증번호']).trim() === String(row['영수증번호']).trim();
                });
                if (포스결제) {
                    카드사명 = 포스결제['매입사'] || '';
                }
            }

            // 정산일자 계산 (T+2)
            const 결제일 = row['일자'];
            let 정산일자 = '';
            if (결제일 && 카드결제) {
                const parts = 결제일.split('-');
                if (parts.length === 3) {
                    const date = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
                    date.setDate(date.getDate() + 2); // T+2
                    정산일자 = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
                }
            }

            stat.품목별상세.push({
                결제일자: row['일자'],
                정산일자: 정산일자,
                카드여부: 카드결제,
                카드사: 카드사명,
                카드수수료율: 카드결제 ? 카드수수료율 : 0,
                상품명: row['상품명'],
                상품코드: 상품코드 || '-',
                상품바코드: 바코드 || '-',
                판매단가: 판매단가,
                수량: parseInt(row['수량']) || 1,
                총매출액: 총매출,
                할인여부: 할인액 > 0,
                할인율: 할인율,
                할인액: 할인액,
                부담주체: 분석.부담주체,
                카드수수료: 카드수수료,
                정산액: 최종정산액
            });
        });

        // 정산액 순 정렬
        const 정렬된결과 = Object.values(거래처별정산)
            .sort((a, b) => b.정산예정액 - a.정산예정액);

        // 요약 출력
        console.log(`품목: 카드 ${카드품목수}건, 비카드 ${비카드품목수}건`);
        console.log(`\n총 ${정렬된결과.length}개 거래처`);
        정렬된결과.slice(0, 10).forEach(stat => {
            console.log(`  ${stat.거래처}: ${stat.품목수}건 (카드 ${stat.카드품목수}/비카드 ${stat.비카드품목수}), 정산 ${stat.정산예정액.toLocaleString()}원`);
        });

        return 거래처별정산;
    }

    // 4. 수수료율 역산 및 검증
    calculate실제수수료율() {
        console.log('\n=== 수수료율 역산 분석 ===');

        const 카드사별통계 = {};

        this.매입내역.forEach(row => {
            const 카드사 = row['카드사'] || row['CARD_COMPANY'];
            const 매입금액 = parseFloat(row['매입금액'] || row['AMOUNT'] || 0);
            const 수수료합계 = parseFloat(row['수수료합계(B)'] || row['FEE_TOTAL'] || 0);

            if (!카드사 || 매입금액 === 0) return;

            // 수수료율을 소수점 첫째자리로 반올림 (2.300%, 2.299%, 2.288% → 2.3%)
            const 실제수수료율_정밀 = (수수료합계 / 매입금액 * 100);
            const 실제수수료율 = 실제수수료율_정밀.toFixed(1); // 소수점 첫째자리

            if (!카드사별통계[카드사]) {
                카드사별통계[카드사] = {
                    건수: 0,
                    총매입금액: 0,
                    총수수료: 0,
                    수수료율분포: {}
                };
            }

            카드사별통계[카드사].건수++;
            카드사별통계[카드사].총매입금액 += 매입금액;
            카드사별통계[카드사].총수수료 += 수수료합계;
            카드사별통계[카드사].수수료율분포[실제수수료율] =
                (카드사별통계[카드사].수수료율분포[실제수수료율] || 0) + 1;
        });

        // 결과 출력 및 검증
        Object.keys(카드사별통계).sort().forEach(카드사 => {
            const stat = 카드사별통계[카드사];
            const 평균수수료율 = (stat.총수수료 / stat.총매입금액 * 100).toFixed(1); // 소수점 첫째자리
            const 예상수수료율 = CARD_FEE_INFO[카드사];

            console.log(`\n${카드사}:`);
            console.log(`  건수: ${stat.건수}건`);
            console.log(`  총 매입금액: ${stat.총매입금액.toLocaleString()}원`);
            console.log(`  총 수수료: ${stat.총수수료.toLocaleString()}원`);
            console.log(`  평균 수수료율: ${평균수수료율}%`);

            if (예상수수료율) {
                const 평균율 = parseFloat(평균수수료율);
                // 부가세 포함 여부 확인 (예: 2.3% = 2.09% × 1.1)
                const 부가세제외율 = (평균율 / 1.1).toFixed(1);

                let matched = null;
                if (Math.abs(평균율 - 예상수수료율.online) < 0.3) {
                    matched = 'online';
                } else if (Math.abs(평균율 - 예상수수료율.expert) < 0.3) {
                    matched = 'expert';
                } else if (Math.abs(parseFloat(부가세제외율) - 예상수수료율.online) < 0.3) {
                    matched = 'online+부가세';
                } else if (Math.abs(parseFloat(부가세제외율) - 예상수수료율.expert) < 0.3) {
                    matched = 'expert+부가세';
                }

                if (matched) {
                    if (matched.includes('부가세')) {
                        const baseType = matched.replace('+부가세', '');
                        console.log(`  ✓ 예상 수수료 유형: ${baseType} (${예상수수료율[baseType]}%) + 부가세 10%`);
                        console.log(`    → 부가세 제외: ${부가세제외율}%`);
                    } else {
                        console.log(`  ✓ 예상 수수료 유형: ${matched} (${예상수수료율[matched]}%)`);
                    }
                } else {
                    console.log(`  ⚠ 예상 수수료율과 불일치! (online: ${예상수수료율.online}%, expert: ${예상수수료율.expert}%)`);
                }
            } else {
                console.log(`  ⚠ 수수료 정보 없음`);
            }

            // 수수료율 분포 (상위 5개) - 소수점 첫째자리로 통합됨
            const 분포 = Object.entries(stat.수수료율분포)
                .sort((a, b) => b[1] - a[1])
                .slice(0, 5);
            console.log(`  수수료율 분포:`);
            분포.forEach(([율, 건수]) => {
                const 비율 = ((건수 / stat.건수) * 100).toFixed(1);
                console.log(`    ${율}%: ${건수}건 (${비율}%)`);
            });
        });

        return 카드사별통계;
    }

    // 4-1. 승인/취소 상쇄 거래 감지 (강화된 버전)
    detect상쇄거래() {
        const 상쇄쌍 = [];
        const 승인목록 = [];
        const 취소목록 = [];

        // 승인/취소 분리
        this.포스결제내역.forEach(row => {
            const 구분 = row['승인_구분'] || row['구분'] || row['승인구분'] || '';
            if (구분 === '승인' || 구분.includes('승인')) {
                승인목록.push(row);
            } else if (구분 === '취소' || 구분.includes('취소')) {
                취소목록.push(row);
            }
        });

        const 사용된승인 = new Set();
        const 사용된취소 = new Set();

        // 방법 1: 승인번호로 매칭 (우선순위 높음) - 1:1 매칭
        승인목록.forEach((승인, 승인idx) => {
            if (사용된승인.has(승인idx)) return;

            const 승인번호 = String(승인.승인번호 || '').trim();
            if (!승인번호) return;

            // 아직 사용 안 된 취소 중에서 첫 번째 매칭만 찾기
            for (let 취소idx = 0; 취소idx < 취소목록.length; 취소idx++) {
                if (사용된취소.has(취소idx)) continue;

                const 취소 = 취소목록[취소idx];
                const 취소승인번호 = String(취소.승인번호 || '').trim();
                const 승인금액 = Math.abs(parseFloat(승인.승인금액 || 0));
                const 취소금액 = Math.abs(parseFloat(취소.승인금액 || 0));

                // 카드번호 비교 추가 (같은 카드에서만 상쇄 가능)
                const 승인카드번호 = String(승인.카드번호 || '').trim();
                const 취소카드번호 = String(취소.카드번호 || '').trim();

                // 승인번호 절대값 비교 (취소는 -가 붙음) + 취소일자 >= 승인일자 + 7일 이내 + 같은 카드
                const 승인일자 = 승인.영업일자 || 0;
                const 취소일자 = 취소.영업일자 || 0;
                const 날짜차이 = 취소일자 - 승인일자;

                if (Math.abs(parseInt(취소승인번호)) === Math.abs(parseInt(승인번호)) &&
                    승인금액 === 취소금액 &&
                    승인카드번호 === 취소카드번호 &&
                    날짜차이 >= 0 && 날짜차이 <= 7) {
                    상쇄쌍.push({
                        승인: 승인,
                        취소: 취소,
                        금액: 승인금액,
                        카드사: 승인.매입사 || '',
                        승인일자: 승인.영업일자 || '',
                        취소일자: 취소.영업일자 || '',
                        매칭방법: '승인번호+금액',
                        출처: '포스'
                    });
                    사용된승인.add(승인idx);
                    사용된취소.add(취소idx);
                    break; // 1:1 매칭이므로 하나 찾으면 끝
                }
            }
        });

        // 방법 2: 승인번호 없거나 방법1에서 매칭 안 된 경우 - 카드번호 + 금액 + 날짜로 매칭
        승인목록.forEach((승인, 승인idx) => {
            // 이미 매칭된 경우 스킵
            if (사용된승인.has(승인idx)) return;

            const 승인금액 = Math.abs(parseFloat(승인.승인금액 || 0));
            const 승인카드번호 = String(승인.카드번호 || '').trim();
            const 승인카드사 = 승인.매입사 || '';
            const 승인일자 = 승인.영업일자 || 0;

            if (승인금액 === 0) return;

            for (let 취소idx = 0; 취소idx < 취소목록.length; 취소idx++) {
                if (사용된취소.has(취소idx)) continue;

                const 취소 = 취소목록[취소idx];
                const 취소금액 = Math.abs(parseFloat(취소.승인금액 || 0));
                const 취소카드번호 = String(취소.카드번호 || '').trim();
                const 취소카드사 = 취소.매입사 || '';
                const 취소일자 = 취소.영업일자 || 0;

                const 날짜차이 = 취소일자 - 승인일자;

                // 카드번호 + 금액 + 날짜 + 카드사 (취소는 승인 이후)
                if (승인카드번호 && 취소카드번호 &&
                    승인카드번호 === 취소카드번호 &&
                    승인금액 === 취소금액 &&
                    승인카드사 === 취소카드사 &&
                    날짜차이 >= 0 && 날짜차이 <= 7) {

                    상쇄쌍.push({
                        승인: 승인,
                        취소: 취소,
                        금액: 승인금액,
                        카드사: 승인카드사,
                        승인일자: 승인일자,
                        취소일자: 취소일자,
                        날짜차이: 날짜차이,
                        매칭방법: '카드번호+금액+날짜',
                        출처: '포스'
                    });
                    사용된승인.add(승인idx);
                    사용된취소.add(취소idx);
                    break;
                }
            }
        });

        // 방법 3: 카드사 + 금액 + 같은 날짜
        승인목록.forEach((승인, 승인idx) => {
            if (사용된승인.has(승인idx)) return;

            const 승인금액 = Math.abs(parseFloat(승인.승인금액 || 0));
            const 승인카드사 = 승인.매입사 || '';
            const 승인일자 = 승인.영업일자 || 0;

            if (승인금액 === 0) return;

            for (let 취소idx = 0; 취소idx < 취소목록.length; 취소idx++) {
                if (사용된취소.has(취소idx)) continue;

                const 취소 = 취소목록[취소idx];
                const 취소금액 = Math.abs(parseFloat(취소.승인금액 || 0));
                const 취소카드사 = 취소.매입사 || '';
                const 취소일자 = 취소.영업일자 || 0;

                if (승인카드사 === 취소카드사 &&
                    승인금액 === 취소금액 &&
                    승인일자 === 취소일자) {

                    상쇄쌍.push({
                        승인: 승인,
                        취소: 취소,
                        금액: 승인금액,
                        카드사: 승인카드사,
                        승인일자: 승인일자,
                        취소일자: 취소일자,
                        날짜차이: 0,
                        매칭방법: '카드사+금액+날짜',
                        출처: '포스'
                    });
                    사용된승인.add(승인idx);
                    사용된취소.add(취소idx);
                    break;
                }
            }
        });

        return 상쇄쌍;
    }

    // 4-2. 매입내역에서 상쇄 거래 감지
    detect매입상쇄거래() {
        const 상쇄쌍 = [];
        const 승인목록 = [];
        const 취소목록 = [];

        // 매입내역에서 승인/취소 분리 (금액의 부호로 구분)
        this.매입내역.forEach(row => {
            const 매입금액 = parseFloat(row.매입금액 || 0);

            if (매입금액 > 0) {
                승인목록.push(row);
            } else if (매입금액 < 0) {
                취소목록.push(row);
            }
        });

        const 사용된취소 = new Set();

        // 방법 1: 승인번호로 매칭
        승인목록.forEach(승인 => {
            const 승인번호 = String(승인.승인번호 || '').trim();
            if (!승인번호) return;

            const 매칭취소들 = 취소목록.filter(취소 => {
                const 취소승인번호 = String(취소.승인번호 || '').trim();
                const 승인금액 = Math.abs(parseFloat(승인.매입금액 || 0));
                const 취소금액 = Math.abs(parseFloat(취소.매입금액 || 0));

                return 취소승인번호 === 승인번호 && 승인금액 === 취소금액;
            });

            매칭취소들.forEach(취소 => {
                const 취소idx = 취소목록.indexOf(취소);
                if (!사용된취소.has(취소idx)) {
                    상쇄쌍.push({
                        승인: 승인,
                        취소: 취소,
                        금액: Math.abs(parseFloat(승인.매입금액 || 0)),
                        카드사: 승인.카드사 || '',
                        승인일자: 승인.거래일자 || '',
                        취소일자: 취소.거래일자 || '',
                        매칭방법: '승인번호+금액',
                        출처: '매입'
                    });
                    사용된취소.add(취소idx);
                }
            });
        });

        // 방법 2: 카드번호 + 금액 + 날짜 근접도
        승인목록.forEach(승인 => {
            const 승인번호 = String(승인.승인번호 || '').trim();

            if (승인번호 && 상쇄쌍.some(pair => pair.승인 === 승인)) {
                return;
            }

            const 승인금액 = Math.abs(parseFloat(승인.매입금액 || 0));
            const 승인카드번호 = String(승인.카드번호 || '').trim();
            const 승인일자 = 승인.거래일자 || '';

            if (!승인카드번호 || 승인금액 === 0) return;

            취소목록.forEach((취소, idx) => {
                if (사용된취소.has(idx)) return;

                const 취소금액 = Math.abs(parseFloat(취소.매입금액 || 0));
                const 취소카드번호 = String(취소.카드번호 || '').trim();
                const 취소일자 = 취소.거래일자 || '';

                // 날짜 문자열 차이 계산 (YYYYMMDD 형식 가정)
                const 날짜차이 = Math.abs(parseInt(승인일자) - parseInt(취소일자));

                if (승인카드번호 === 취소카드번호 &&
                    승인금액 === 취소금액 &&
                    날짜차이 <= 7) {

                    상쇄쌍.push({
                        승인: 승인,
                        취소: 취소,
                        금액: 승인금액,
                        카드사: 승인.카드사 || '',
                        승인일자: 승인일자,
                        취소일자: 취소일자,
                        날짜차이: 날짜차이,
                        매칭방법: '카드번호+금액+날짜',
                        출처: '매입'
                    });
                    사용된취소.add(idx);
                }
            });
        });

        return 상쇄쌍;
    }

    // 5. 매입내역과 포스결제내역 매칭
    match매입포스() {
        console.log('\n=== 매입내역 ↔ 포스결제내역 매칭 ===');

        const 매칭된항목 = [];
        const 매입누락 = [];
        const 포스누락 = [];

        // 1단계: 포스 내부 상쇄 거래 감지 (매입은 이미 정산된 최종 데이터)
        const 포스상쇄거래_원본 = this.detect상쇄거래();

        console.log(`\n[포스 상쇄 거래 감지]`);
        console.log(`  포스 상쇄 후보: ${포스상쇄거래_원본.length}쌍`);

        // ★ 핵심 수정: 매입 파일에 해당 거래가 있으면 상쇄에서 제외
        // (매입에 있다 = 정산됨 = 실제로 취소되지 않은 거래)
        const excelDateToYYYYMMDD_temp = (serial) => {
            if (!serial) return '';
            if (typeof serial === 'string' && /^\d{8}$/.test(serial)) return serial;
            if (typeof serial === 'number') {
                const epoch = new Date(1899, 11, 30);
                const date = new Date(epoch.getTime() + serial * 86400000);
                const year = date.getFullYear();
                const month = String(date.getMonth() + 1).padStart(2, '0');
                const day = String(date.getDate()).padStart(2, '0');
                return `${year}${month}${day}`;
            }
            return String(serial);
        };

        // 매입 데이터로 빠른 검색을 위한 인덱스 구축
        const 매입인덱스 = new Set();
        this.매입내역.forEach(row => {
            const 거래일자 = String(row.거래일자 || '').replace(/-/g, '');
            const 매입금액 = parseFloat(row.매입금액 || 0);
            const 카드사 = row.카드사 || '';
            if (거래일자 && 매입금액 > 0) {
                매입인덱스.add(`${거래일자}_${매입금액}_${카드사}`);
            }
        });

        // 상쇄 쌍 검증: 승인건이 매입에 있으면 실제 상쇄가 아님
        const 포스상쇄거래 = 포스상쇄거래_원본.filter(pair => {
            const 승인날짜 = excelDateToYYYYMMDD_temp(pair.승인.영업일자);
            const 승인금액 = Math.abs(parseFloat(pair.승인.승인금액 || 0));
            const 승인카드사 = pair.승인.매입사 || '';
            const key = `${승인날짜}_${승인금액}_${승인카드사}`;

            if (매입인덱스.has(key)) {
                // 매입에 있음 = 정산된 거래 = 실제 취소 아님
                console.log(`  ⚠ 상쇄 제외: No.${pair.승인['No.']} (매입에 존재 - 정산된 거래)`);
                return false;
            }
            return true;
        });

        console.log(`  포스 유효 상쇄: ${포스상쇄거래.length}쌍 (매입 검증 후)`);

        const 모든상쇄거래 = 포스상쇄거래;

        // 2단계: 포스결제내역에서 승인건만 필터링 (상쇄된 거래 제외)
        // ★ 중요: 승인번호로 필터하면 안 됨! 같은 승인번호가 다른 거래에 나타날 수 있음
        //          (예: No.477 롯데카드 9000, No.482 KB카드 9000)
        //          → No. 값으로 정확히 필터해야 함
        const 포스상쇄된NoSet = new Set(포스상쇄거래.flatMap(pair => [
            pair.승인['No.'],
            pair.취소['No.']
        ]));

        const 포스승인건 = this.포스결제내역.filter(row => {
            const 구분 = row['승인_구분'] || row['구분'] || row['승인구분'] || '';
            const rowNo = row['No.'];

            // 승인건이면서 상쇄되지 않은 거래만
            return (구분 === '승인' || 구분.includes('승인')) && !포스상쇄된NoSet.has(rowNo);
        });

        console.log(`포스 승인건: ${포스승인건.length}건 (상쇄 제외 후)`);

        // 3단계: 매입내역은 양수 금액만 (이미 정산된 최종 데이터이므로 상쇄 처리 불필요)
        const 매입유효건 = this.매입내역.filter(row => {
            const 매입금액 = parseFloat(row.매입금액 || 0);
            return 매입금액 > 0;
        });

        console.log(`매입 유효건: ${매입유효건.length}건`);

        // ★ 중요: 포스와 매입의 승인번호 체계가 완전히 다름!
        //    포스: 4-5자리 (9000, 5454, 181 등)
        //    매입: 8자리 (30035073, 79413800 등)
        // → 승인번호로 매칭 불가능! 날짜+금액+카드사로 매칭해야 함

        console.log(`\n[매칭 방식: 날짜+금액+카드사 우선]`);

        // Excel 날짜 serial을 YYYYMMDD로 변환하는 함수
        const excelDateToYYYYMMDD = (serial) => {
            if (!serial) return '';

            // 이미 YYYYMMDD 형식이면 그대로 반환
            if (typeof serial === 'string' && /^\d{8}$/.test(serial)) {
                return serial;
            }

            // Excel serial number인 경우 변환
            if (typeof serial === 'number') {
                const epoch = new Date(1899, 11, 30);
                const date = new Date(epoch.getTime() + serial * 86400000);
                const year = date.getFullYear();
                const month = String(date.getMonth() + 1).padStart(2, '0');
                const day = String(date.getDate()).padStart(2, '0');
                return `${year}${month}${day}`;
            }

            return String(serial);
        };

        // 매입내역을 날짜+금액+카드사로 인덱싱
        const 매입Map = new Map();
        매입유효건.forEach(매입row => {
            const 거래일자_raw = String(매입row.거래일자 || '').trim();
            const 거래일자 = 거래일자_raw.replace(/-/g, '');  // YYYY-MM-DD → YYYYMMDD
            const 매입금액 = parseFloat(매입row.매입금액 || 0);
            const 카드사 = 매입row.카드사 || '';

            if (거래일자 && 매입금액 > 0) {
                // 우선순위 1: 날짜+금액+카드사
                const key1 = `${거래일자}_${매입금액}_${카드사}`;
                if (!매입Map.has(key1)) {
                    매입Map.set(key1, []);
                }
                매입Map.get(key1).push({ row: 매입row, key: key1, method: '날짜+금액+카드사' });

                // 우선순위 2: 날짜+금액 (카드사 매칭 실패 시 사용)
                const key2 = `${거래일자}_${매입금액}`;
                if (!매입Map.has(key2)) {
                    매입Map.set(key2, []);
                }
                매입Map.get(key2).push({ row: 매입row, key: key2, method: '날짜+금액' });
            }
        });

        // 포스 내역을 날짜+금액+카드사로 매칭
        포스승인건.forEach(포스row => {
            const 영업일자 = excelDateToYYYYMMDD(포스row['영업일자'] || 포스row['승인일자']);
            const 승인금액 = parseFloat(포스row['승인금액'] || 포스row['승인요청금액'] || 0);
            const 카드사 = 포스row['매입사'] || '';

            if (!영업일자 || 승인금액 <= 0) {
                포스누락.push({
                    이유: '날짜/금액 없음',
                    승인번호: 포스row.승인번호,
                    포스데이터: 포스row
                });
                return;
            }

            // 우선순위 1: 날짜+금액+카드사 매칭
            let key = `${영업일자}_${승인금액}_${카드사}`;
            let 매칭된매입 = 매입Map.get(key);
            let 매칭방법 = '날짜+금액+카드사';

            // 우선순위 2: 날짜+금액만 매칭
            if (!매칭된매입 || 매칭된매입.length === 0) {
                key = `${영업일자}_${승인금액}`;
                매칭된매입 = 매입Map.get(key);
                매칭방법 = '날짜+금액';
            }

            if (매칭된매입 && 매칭된매입.length > 0) {
                const 매입데이터 = 매칭된매입[0];
                매칭된항목.push({
                    포스: 포스row,
                    매입: 매입데이터.row,
                    매칭방법: 매칭방법,
                    포스승인번호: 포스row.승인번호,
                    매입승인번호: 매입데이터.row.승인번호
                });

                // 매칭된 항목 제거 (1:1 매칭)
                매칭된매입.shift();
                if (매칭된매입.length === 0) {
                    매입Map.delete(key);
                }
            } else {
                포스누락.push({
                    이유: '매입내역 없음',
                    날짜: 영업일자,
                    금액: 승인금액,
                    카드사: 카드사,
                    승인번호: 포스row.승인번호,
                    포스데이터: 포스row
                });
            }
        });

        // 매칭되지 않은 매입내역 처리
        const 처리된매입 = new Set();
        매칭된항목.forEach(item => {
            처리된매입.add(item.매입);
        });

        매입유효건.forEach(매입row => {
            if (!처리된매입.has(매입row)) {
                매입누락.push({
                    이유: '포스내역 없음',
                    매입승인번호: 매입row.승인번호,
                    매입데이터: 매입row
                });
            }
        });

        console.log(`✓ 매칭 성공: ${매칭된항목.length}건`);
        console.log(`⚠ 포스에만 있음: ${포스누락.length}건`);
        console.log(`⚠ 매입에만 있음: ${매입누락.length}건`);

        this.검증결과.매칭검증 = {
            매칭된항목,
            포스누락,
            매입누락,
            상쇄된거래: 모든상쇄거래
        };

        return this.검증결과.매칭검증;
    }

    // 6. 영수증과 카드결제 연결
    match영수증카드() {
        console.log('\n=== 영수증 ↔ 카드결제 연결 ===');

        // 영수증별 합계 계산
        const 영수증합계 = {};
        this.영수증내역.forEach(row => {
            const key = `${row['일자']}_${row['영수증번호']}`;
            if (!영수증합계[key]) {
                영수증합계[key] = {
                    일자: row['일자'],
                    영수증번호: row['영수증번호'],
                    결제시각: row['결제시각'],
                    총금액: 0,
                    품목수: 0,
                    품목들: []
                };
            }
            영수증합계[key].총금액 += parseFloat(row['실매출액'] || 0);
            영수증합계[key].품목수++;
            영수증합계[key].품목들.push(row['상품명']);
        });

        console.log(`영수증 개수: ${Object.keys(영수증합계).length}개`);
        console.log(`총 품목 수: ${this.영수증내역.length}개`);

        // 샘플 출력
        const 샘플 = Object.values(영수증합계).slice(0, 5);
        console.log('\n샘플 영수증:');
        샘플.forEach(영수증 => {
            console.log(`  ${영수증.일자} ${영수증.영수증번호}: ${영수증.총금액.toLocaleString()}원 (${영수증.품목수}개 품목)`);
        });

        return 영수증합계;
    }

    // 7. 전체 리포트 생성
    generateReport() {
        const report = {
            생성시각: new Date().toISOString(),
            요약: {
                매입내역건수: this.매입내역.length,
                포스결제건수: this.포스결제내역.length,
                영수증품목수: this.영수증내역.length
            },
            검증결과: this.검증결과
        };

        const filename = `검증리포트_${new Date().toISOString().split('T')[0]}.json`;
        fs.writeFileSync(filename, JSON.stringify(report, null, 2), 'utf-8');
        console.log(`\n✓ 리포트 저장: ${filename}`);

        return report;
    }
}

// 실행
if (require.main === module) {
    console.log('=== 카드 결제 데이터 분석 시작 ===\n');

    const analyzer = new CardDataAnalyzer();

    try {
        // 데이터 로드
        analyzer.load매입내역();
        analyzer.load포스결제내역();
        analyzer.load영수증내역();

        // 분석 실행
        analyzer.calculate실제수수료율();
        analyzer.match매입포스();
        analyzer.match영수증카드();

        // 리포트 생성
        analyzer.generateReport();

        console.log('\n=== 분석 완료 ===');
    } catch (error) {
        console.error('에러 발생:', error.message);
        console.error(error.stack);
    }
}

module.exports = CardDataAnalyzer;
