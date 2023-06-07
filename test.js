import * as ExcelJS from 'exceljs';
import { IMovieSetting } from '../interface/property';
import { IExcelResponse } from '../interface/reponse/resposne';
import { IMovieCrawlRequest } from '../interface/request/movie-crawl.request';
import path from 'path';
import _ from 'lodash';
import { MovieNameEnum } from '../util/crawl/interface/crawl';
import { ExcelUtil } from '../util/excelUtil';
import { Injectable } from '@nestjs/common';
import { ObjectUtil } from '../util/objectUtil';
import { IheaderInfo } from '../util/crawl/interface/excel';
import { logger } from '../util/crawl/config/logger/logger';
import { getAreaNmFromRefine } from '../api/convert_name/default';

/**
 * @description: ExcelUtil / ObjectUtil 클래스 안의 static 라이브러리 함수들을 편하게 사용하기 위해 커스텀으로 직접 구현하여 사용중
 */
const { setWorkSheet, setHeader, setBorder, setFill, setCenter, setBold, setFont2, setBorder2, setRight } = ExcelUtil;
const { numberToColumn, newJson } = ObjectUtil;

interface IScreenTabeData {
  theaterCompany: string; // 계열사
  areaNm: string; // 지역명
  theaterNm: string; // 극장명
  note: string; // 자막 / 더빙 구분
  type: string; // 2D or 3D 등등
  totalSeat: number; // 좌석수
  hall: string; // 상영관
  totalTimeCnt: number; // 총 회차
  screenCntCnt: number; // 총 상영수 ( 일부 코드는 해당 값에 상영관 수로 지정하기도 했다. )
  totalSeatCnt: number; // 총 좌석수
  seatRateAvg: number; // 좌판율
  timeLen: number; // 회차 길이 ( totalTimeCnt와 같은 값을 가진 것으로 기억 )
  soldSeatCnt: number; // 판매된 좌석 수
  totalRemainSeat: number; // 총 남은 좌석 수
  idx: string | number; // index
}

type TScreenTable = Array<{
  screenDt: string;
  data: {
    data: {
      [theaterCompany: string]: Array<IScreenTabeData>;
    };
    totalTimeLength: any;
  };
}>;

interface IDataByTheaterCompanyInfo {
  theaterCnt: number; // 극장수
  timeCnt: number; // 상영회차
  screenCnt: number; // 상영관 수
  totalSeatCnt: number; // 총 좌석수
  timeAvg: number; // 평균회차
  seatCntAvg: number; // 평균좌석수
  seatRateAvg: number | string; // 평균 좌판율
  totalRemainSeat?: number; // 총 남은 좌석 수
  hallCnt?: number; // 상영관 수
}

@Injectable()
export class ExcelService {
  /**
   * @description: 배열 데이터 정렬 함수 / 지역별 -> 상영관 별
   */
  sortRefinedData(arr: Array<any>): Array<any> {
    return arr.sort((data1: any, data2: any) => {
      if (data1.areaNm < data2.areaNm) {
        return -1;
      } else if (data1.areaNm > data2.areaNm) {
        return 1;
      } else {
        if (data1.theaterNm < data2.theaterNm) {
          return -1;
        } else {
          return 1;
        }
      }
    });
  }

  /**
   * @description: 배열 형태의 데이터를 계열사별 시트 데이터에 맞게 변환
   */
  createDataByTheaterCompany(
    mainMovieData: Array<IMovieCrawlRequest>,
    screenDt: string
  ): {
    [screenDt: string]: { [theaterCompany: string]: IDataByTheaterCompanyInfo };
  } {
    const theaterNmListByTheaterCompany = {};
    const testResult = {};

    let lastHallNm = '';

    const refineData = this.createScreenTableData(mainMovieData, []).data;
    Object.keys(refineData).forEach((theaterCompany) => {
      if (!testResult[theaterCompany]) {
        testResult[theaterCompany] = {
          seatCntAvg: 0,
          seatRateAvg: 0,
          timeAvg: 0,
          screenCnt: 0,
          theaterCnt: 0,
          timeCnt: 0,
          totalSeatCnt: 0,
          totalRemainSeat: 0,
          hallCnt: 0,
        };
      }

      if (!theaterNmListByTheaterCompany[theaterCompany]) {
        theaterNmListByTheaterCompany[theaterCompany] = [];
      }

      const value = refineData[theaterCompany];
      value.forEach((d) => {
        const { totalTimeCnt, totalRemainSeat, totalSeatCnt, theaterNm, hall, theaterCompany } = d;

        testResult[theaterCompany] = {
          ...testResult[theaterCompany],
          timeCnt: testResult[theaterCompany].timeCnt + totalTimeCnt,
          screenCnt: 0,
          theaterCnt: 0,
          totalSeatCnt: testResult[theaterCompany].totalSeatCnt + totalSeatCnt,
          totalRemainSeat:
            testResult[theaterCompany].totalRemainSeat + (totalRemainSeat === -1 ? totalSeatCnt : totalRemainSeat),
        };

        const hallNm = theaterNm + hall;
        if (lastHallNm.length === 0 || hallNm !== lastHallNm) {
          testResult[theaterCompany].hallCnt = testResult[theaterCompany].hallCnt += 1;
        }

        lastHallNm = hallNm;

        if (!theaterNmListByTheaterCompany[theaterCompany].includes(theaterNm)) {
          theaterNmListByTheaterCompany[theaterCompany].push(theaterNm);
        }
      });

      lastHallNm = '';
    });

    Object.keys(testResult).forEach((objTheaterCompany) => {
      const dd = testResult[objTheaterCompany];
      const soldSeatCnt = dd.totalSeatCnt - dd.totalRemainSeat;

      testResult[objTheaterCompany] = {
        ...testResult[objTheaterCompany],
        theaterCnt: theaterNmListByTheaterCompany[objTheaterCompany].length,
        seatCntAvg: dd.totalSeatCnt / dd.timeCnt,
        timeAvg: dd.timeCnt / theaterNmListByTheaterCompany[objTheaterCompany].length,
        seatRateAvg: ((soldSeatCnt / dd.totalSeatCnt) * 100).toFixed(1),
      };
    });

    return { [screenDt]: testResult };
  }

  /**
   * @description: 데이터 정렬 함수 / object의 key 값을 배열로 변환 후 오름차순 정렬
   */
  sortObjKeyAsc(obj: any) {
    return Object.keys(obj).sort((a, b) => (a < b ? -1 : 1));
  }

  /**
   * @description: 배열 형태의 데이터를 지역별 시트 데이터 형식에 맞춰 변환
   */
  createDataByArea(
    mainMovieData: Array<IMovieCrawlRequest>,
    screenDt: string
  ): {
    [screenDt: string]: { [areaNm: string]: IDataByTheaterCompanyInfo };
  } {
    const theaterNmListByTheaterCompany = {};

    const refineData = this.createScreenTableData(mainMovieData, []).data;
    let lastHallNm = '';
    const data = Object.keys(refineData).reduce((acc, theaterCompany) => {
      const value = refineData[theaterCompany];
      value.forEach((d) => {
        const { totalTimeCnt, totalRemainSeat, totalSeatCnt, theaterNm, hall, theaterCompany, areaNm } = d;
        if (!acc[areaNm]) {
          acc[areaNm] = {
            seatCntAvg: 0,
            seatRateAvg: 0,
            timeAvg: 0,
            screenCnt: 0,
            theaterCnt: 0,
            timeCnt: 0,
            totalSeatCnt: 0,
            totalRemainSeat: 0,
            hallCnt: 0,
          };
        }

        if (!theaterNmListByTheaterCompany[areaNm]) {
          theaterNmListByTheaterCompany[areaNm] = [];
        }

        acc[areaNm] = {
          ...acc[areaNm],
          timeCnt: acc[areaNm].timeCnt + totalTimeCnt,
          screenCnt: 0,
          theaterCnt: 0,
          totalSeatCnt: acc[areaNm].totalSeatCnt + totalSeatCnt,
          totalRemainSeat: acc[areaNm].totalRemainSeat + (totalRemainSeat === -1 ? totalSeatCnt : totalRemainSeat),
        };

        const hallNm = screenDt + theaterCompany + theaterNm + hall;

        if (lastHallNm.length === 0 || hallNm !== lastHallNm) {
          acc[areaNm].hallCnt = acc[areaNm].hallCnt += 1;
        }

        lastHallNm = hallNm;

        if (!theaterNmListByTheaterCompany[areaNm].includes(theaterCompany + theaterNm)) {
          theaterNmListByTheaterCompany[areaNm].push(theaterCompany + theaterNm);
        }
      });

      return acc;
    }, {});

    Object.keys(data).forEach((subAreaNm) => {
      const dd = data[subAreaNm];
      const soldSeatCnt = dd.totalSeatCnt - dd.totalRemainSeat;

      data[subAreaNm] = {
        ...data[subAreaNm],
        theaterCnt: theaterNmListByTheaterCompany[subAreaNm].length,
        seatCntAvg: dd.totalSeatCnt / dd.timeCnt,
        timeAvg: dd.timeCnt / theaterNmListByTheaterCompany[subAreaNm].length,
        seatRateAvg: ((soldSeatCnt / dd.totalSeatCnt) * 100).toFixed(1),
      };
    });
    return { [screenDt]: data };
  }

  /**
   * @description: 극장명 앞에 계열사 명이 있는지 체크 후, 없으면 계열사 + 극장명으로 반환
   */
  refineTheaterNm(teaterCompany, theaterNm) {
    switch (teaterCompany) {
      case 'cgv':
        return /^CGV/g.test(theaterNm) || /^cgv/g.test(theaterNm) ? theaterNm : 'CGV' + theaterNm;
      case 'mega':
        return /^메가박스/g.test(theaterNm) ? theaterNm : '메가박스' + theaterNm;
      case 'lotte':
        return /^LOTTE/g.test(theaterNm) || /^lotte/g.test(theaterNm) || /^롯데/g.test(theaterNm)
          ? theaterNm
          : '롯데' + theaterNm;
    }
  }

  /**
   * @description: 이전 데이터와, 현재 데이터를 비교해서 포맷(자막, 더빙) 을 제외하고 나머지가 일치하는지 확인
   */
  isSameData(row1, row2) {
    return (
      row1.getCell('type').value === row2.getCell('type').value &&
      row1.getCell('totalSeat').value === row2.getCell('totalSeat').value &&
      row1.getCell('hall').value === row2.getCell('hall').value &&
      row1.getCell('theaterNm').value === row2.getCell('theaterNm').value
    );
  }

  /**
   * @description: 해당 행의 숫자를 입력하는 컬럼에서, 숫자 포맷으로 변경
   */
  getNumFmtRow(row: ExcelJS.Row, targetCellNmList: Array<string>) {
    targetCellNmList.forEach((cellNm) => {
      const cell = row.getCell(cellNm);

      if (cellNm === 'totalSeatCnt') {
        row.getCell(cellNm).numFmt = '#,##0석';
      } else {
        row.getCell(cellNm).numFmt = '#,##0';
      }
    });

    return row;
  }

  /**
   * @description: 상영시간표 시트를 만드는 함수
   *
   * @process:
   * 1. 엑셀 시트의 헤더값 설정 및 엑셀 시트 생성
   * 2. 데이터 정렬 및 그룹화 ( 지역-상영관별 )
   * 3. 그룹화된 데이터를 가지고 엑셀 시트에 데이터 추가
   * 3-1. 그룹화 된 상태에서, 반복문을 돌리는데, 그룹화 된 데이터들의 경우 merge 작업 진행
   * 4. excel 스타일링
   */
  async createScreenTableSheet(
    workbook: ExcelJS.Workbook,
    mainMovieNm: string,
    screenDt: string,
    screenTableData: { data: any; totalTimeLength: any },
    crawlTime: string
  ) {
    // 회차 컬럼 설정 ( 초기에는 최대값 기준으로 만들게 했으나, 고정값으로 해달라고 요청하여 12회차까지 나오도록 설정 )
    const timeHeaderList = Array(12)
      .fill(1)
      .map((d, idx) => ({
        header: `${idx + 1}회`,
        key: `time_${idx + 1}`,
        columnKey: `time_${idx + 1}`,
        width: 5,
      }));

    // 시트의 헤더값 설정 ( 아무 설정도 하지 않으면 1 번째 줄로 고정되서 나온다. )
    const header = [
      { header: '지역', key: 'areaNm', columnKey: 'areaNm', width: 5 },
      {
        header: '극장명',
        key: 'theaterNm',
        columnKey: 'theaterNm',
        width: 30,
      },
      { header: '포맷', key: 'type', columnKey: 'type', width: 5 },
      { header: '구분', key: 'note', columnKey: 'note' },
      { header: '관', key: 'hall', columnKey: 'hall' },
      { header: '좌석수', key: 'totalSeat', columnKey: 'totalSeat' },
      ...timeHeaderList,
      { header: '총회차', key: 'totalTimeCnt', columnKey: 'totalTimeCnt' },
      {
        header: '총스크린',
        key: 'screenCntCnt',
        columnKey: 'totalScreenCnt',
      },
      { header: '총좌석수', key: 'totalSeatCnt', columnKey: 'totalSeatCnt' },
      { header: '판매좌석수', key: 'soldSeatCnt', columnKey: 'soldSeatCnt' },
      { header: '좌석판매율', key: 'seatRateAvg', columnKey: 'seatRateAvg' },
    ];

    // 시트 생성
    const workSheet: ExcelJS.Worksheet = setWorkSheet(workbook, '상영시간표_' + screenDt);

    const setHeaderInfo: IheaderInfo = {
      movieNm: mainMovieNm,
      date: `${screenDt} | ${crawlTime}`,
      headDownCnt: 3,
    };
    const { headDownCnt } = setHeaderInfo;
    const defaultLen = 6 + timeHeaderList.length;

    // 상영시간표 헤더 설정
    await setHeader(header, workSheet, setHeaderInfo);

    // E column의 width 값을 15로 설정
    workSheet.getColumn('E').width = 16;

    // 회차 컬럼 다음에 있는 컬럼의 width 값을 15로 지정 ( 총좌석수 )
    workSheet.getColumn(numberToColumn(defaultLen + 3)).width = 15;
    // 회차 컬럼 다음에 있는 컬럼의 width 값을 15로 지정 ( 판매좌석수 )
    workSheet.getColumn(numberToColumn(defaultLen + 4)).width = 15;

    const headerPosY = headDownCnt + 1;
    let addIdx = headerPosY;
    let sum_theaterCnt = 0;
    let sum_totalTimeCnt = 0;
    let sum_screenCntCnt = 0;
    let sum_totalSeatCnt = 0;
    let sum_soldSeatCnt = 0;
    let sum_seatRateAvg = '0';

    let coundTargetIdx = 5;
    let coundRangeArray = [];
    this.sortObjKeyAsc(screenTableData.data).forEach((key) => {
      key = key as MovieNameEnum;

      let theaterCnt = 0;
      let totalTimeCnt = 0;
      let screenCntCnt = 0;
      let totalSeatCnt = 0;
      let soldSeatCnt = 0;
      let seatRateAvg = '0';

      addIdx += 1;

      // 먼저 데이터를 정렬해준다.
      const sortedData = this.sortRefinedData(screenTableData.data[key]);

      // 그리고 정렬된 데이터를 지역 / 상영관 별로 그룹화 해준다. ( 같은 지역/ 같은 상영관 끼리 그룹화 )
      const groupedData = _.groupBy(sortedData, (item) => `${item.areaNm}-${item.theaterNm}`);

      let testIdx = 0;
      Object.values(groupedData).forEach((group: Array<any>) => {
        theaterCnt += 1;
        let lastRow: ExcelJS.Row | null = null;
        const startRow = sortedData.findIndex((d: any) => d.idx === group[0].idx);
        const endRow = sortedData.findIndex((d: any) => d.idx === group[group.length - 1].idx);

        // startRow / endRow가 다르다면, 그룹화된 데이터가 여러개 있다는 뜻으로 간주한다.
        if (startRow !== endRow) {
          // 병합할 셀 지정.
          const mergeRange = `B${startRow + addIdx}:B${endRow + addIdx}`;

          const rows = [];
          group.forEach((d) => {
            d.theaterNm = this.refineTheaterNm(d.theaterCompany, d.theaterNm);
            let row = workSheet.addRow({
              ...d,
              type: this.refineType(d.type),
              theaterCnt: d.theaterCnt,
              totalTimeCnt: d.totalTimeCnt,
              screenCntCnt: d.screenCntCnt,
              totalSeatCnt: d.totalSeatCnt,
              soldSeatCnt: (d.soldSeatCnt ?? 0) + '석',
              seatRateAvg: (d.seatRateAvg ?? '0') + '%',
            });
            row = this.getNumFmtRow(row, ['totalSeat', 'totalTimeCnt', 'screenCntCnt', 'totalSeatCnt']);
            if (lastRow) {
              // 같은 데이터일 경우 병합 진행 ( 스크린수 )
              if (this.isSameData(lastRow, row)) {
                workSheet.mergeCells(`T${lastRow.number}:T${row.number}`);
              }
            }

            lastRow = row;
            rows.push(row);
            totalTimeCnt += d.totalTimeCnt;
            screenCntCnt += d.screenCntCnt;
            totalSeatCnt += d.totalSeatCnt;
            soldSeatCnt += d.soldSeatCnt;
          });

          rows.forEach((row) => {
            this.addEmptyValue(workSheet, row);

            header.forEach(({ key }) => {
              const cell = row.getCell(key);
              const isNotBold = key !== 'theaterNm' && key !== 'hall' && key !== 'totalSeat' && !key.includes('time_');
              setBorder(workSheet, cell, 'thin');
              setCenter(cell);
              setFont2(cell, { size: 9, bold: isNotBold });

              if (key === 'totalSeatCnt' || key === 'soldSeatCnt' || key === 'seatRateAvg') {
                setBorder2(cell, { left: 'medium', right: 'medium' });
              }
            });
          });

          // 병합된 셀 스타일 설정하기
          setBorder(workSheet, `B${startRow + addIdx}`, 'thin');
          workSheet.mergeCells(mergeRange);
        } else {
          // 그룹화된 데이터가 없다고 간주하고, 스크린 수를 제외한, 상영관 등은 병합을 진행하지 않음.
          group[0].theaterNm = this.refineTheaterNm(group[0].theaterCompany, group[0].theaterNm);
          const row = workSheet.addRow({
            ...group[0],
            type: this.refineType(group[0].type),
            theaterCnt: group[0].theaterCnt?.toLocaleString() ?? '',
            totalTimeCnt: group[0].totalTimeCnt?.toLocaleString() ?? '',
            screenCntCnt: group[0].screenCntCnt?.toLocaleString() ?? '',
            totalSeatCnt: (group[0].totalSeatCnt?.toLocaleString() ?? '0') + '석',
            soldSeatCnt: (group[0].soldSeatCnt?.toLocaleString() ?? '0') + '석',
            seatRateAvg: (group[0].seatRateAvg ?? '0') + '%',
          });

          if (lastRow && this.isSameData(lastRow, row)) {
            row.getCell('screenCntCnt').value = null;
            workSheet.mergeCells(`T${lastRow.number}:T${row.number}`);
          }

          lastRow = row;

          totalTimeCnt += group[0].totalTimeCnt;
          screenCntCnt += group[0].screenCntCnt;
          totalSeatCnt += group[0].totalSeatCnt;
          soldSeatCnt += group[0].soldSeatCnt;

          // 특정 열의 셀들의 스타일 지정
          header.forEach(({ key }) => {
            const cell = row.getCell(key);
            const isNotBold = key !== 'theaterNm' && key !== 'hall' && key !== 'totalSeat' && !key.includes('time_');
            setBorder(workSheet, cell, 'thin');
            setFont2(cell, { size: 9, bold: isNotBold });
            setCenter(cell);

            if (key === 'totalSeatCnt' || key === 'soldSeatCnt' || key === 'seatRateAvg') {
              setBorder2(cell, { left: 'medium', right: 'medium' });
            }
          });
        }

        workSheet.getCell(`B${startRow + addIdx}`).alignment = {
          vertical: 'middle',
          horizontal: 'center',
        };

        testIdx = endRow;
      });

      seatRateAvg = ((soldSeatCnt / totalSeatCnt) * 100).toFixed(1);
      addIdx += testIdx;
      addIdx += 1;

      workSheet.addRow({
        areaNm: key.toUpperCase() + '소계',
      });

      workSheet.mergeCells(`B${addIdx}:${numberToColumn(defaultLen)}${addIdx}`);

      const findedRow = this.getNumFmtRow(workSheet.findRow(addIdx), [
        'areaNm',
        'totalSeat',
        'totalTimeCnt',
        'screenCntCnt',
        'totalSeatCnt',
        'soldSeatCnt',
      ]);
      header.forEach((header) => {
        const cell = findedRow.getCell(header.key);

        setCenter(cell);
        setFill(cell, 'ffcccc');
        setFont2(cell, { color: '000000', size: 13, bold: true });
        setBorder2(cell, { left: 'thin', top: 'thick', right: 'thin', bottom: 'thick' });

        const lastRowNumber = findedRow.number - 1;
        switch (header.key) {
          case 'areaNm':
            // @ts-ignore
            cell.value = {
              formula: `COUNTA(B${coundTargetIdx}:B${lastRowNumber})`,
            };
            setBorder2(cell, { left: 'thick' });
            break;
          case 'totalTimeCnt':
            cell.value = totalTimeCnt?.toLocaleString() ?? '';
            break;
          case 'screenCntCnt':
            // @ts-ignore
            cell.value = {
              formula: `COUNTA(T${coundTargetIdx}:T${lastRowNumber})`,
            };
            break;
          case 'totalSeatCnt':
            cell.value = totalSeatCnt?.toLocaleString() ?? '';
            setBorder2(cell, { left: 'medium', right: 'medium' });
            break;
          case 'soldSeatCnt':
            cell.value = soldSeatCnt?.toLocaleString() ?? '';
            setBorder2(cell, { left: 'medium', right: 'medium' });
            break;
          case 'seatRateAvg':
            cell.value = seatRateAvg + '%';
            setBorder2(cell, { left: 'medium', right: 'thick' });
            break;
          case 'type':
            cell.value = key.toUpperCase() + '소계';
            break;
        }
      });
      coundRangeArray.push([coundTargetIdx, findedRow.number - 1]);
      coundTargetIdx = findedRow.number + 1;
      sum_theaterCnt += theaterCnt;
      sum_totalTimeCnt += totalTimeCnt;
      sum_screenCntCnt += screenCntCnt;
      sum_totalSeatCnt += totalSeatCnt;
      sum_soldSeatCnt += soldSeatCnt;
      sum_seatRateAvg = ((soldSeatCnt / totalSeatCnt) * 100).toFixed(1);
    });

    const lastRow = workSheet.addRow({
      areaNm: sum_theaterCnt,
      totalTimeCnt: sum_totalTimeCnt ? sum_totalTimeCnt.toLocaleString() : '0',
      screenCntCnt: '',
      totalSeatCnt: (sum_totalSeatCnt ? sum_totalSeatCnt.toLocaleString() : '0') + '석',
      soldSeatCnt: (sum_soldSeatCnt ? sum_soldSeatCnt.toLocaleString() : '0') + '석',
      seatRateAvg: (((sum_soldSeatCnt / sum_totalSeatCnt) * 100)?.toFixed(1) ?? '0') + '%',
    });
    this.addEmptyValue(workSheet, lastRow);

    workSheet.mergeCells(`B${lastRow.number}:${numberToColumn(defaultLen)}${lastRow.number}`);
    header.forEach(({ key }) => {
      const cell = lastRow.getCell(key);

      setCenter(cell);
      setFill(cell, '0033cc');
      setFont2(cell, { color: 'ffffff', size: 13, bold: true });
      setBorder2(cell, { left: 'thin', top: 'thick', right: 'thin', bottom: 'thick' });

      if (key === 'type') {
        cell.value = '총 계';
        setFont2(cell, { bold: true });
      }
      if (key === 'totalSeatCnt' || key === 'soldSeatCnt' || key === 'seatRateAvg') {
        setBorder2(cell, { left: 'medium', right: 'thick' });
      }

      if (key === 'screenCntCnt') {
        //@ts-ignore
        cell.value = {
          formula: `COUNTA(${this.combineRangeArrayForColumn('T', coundRangeArray)})`,
        };
      }
    });

    const headerRow = workSheet.findRow(4);
    header.forEach(({ key }) => {
      const cell = headerRow.getCell(key);
      setBorder(workSheet, cell, 'thin');
      setCenter(cell);
      if (key === 'totalSeatCnt' || key === 'soldSeatCnt' || key === 'seatRateAvg') {
        setFill(cell, '000000');
        setFont2(cell, { size: 10, color: 'ffffff', bold: true });
      }
    });
  }

  /**
   * @description:
   *  [ [4,10], [11, 16], [17, 20] ] => T4:T10, T11:T16, T17:T20 형태로 변환.
   *  isMulti가 true일 경우 + condition이 "*" 일 경우,
   *  [ [4,10], [11, 16], [17, 20] ] => T4:T10,"*",T11:T16,"*",T17:T20,"*" 형식으로 변환
   */
  combineRangeArrayForColumn(target: string, array: Array<Array<number>>, isMulti = false, condition?: string) {
    return array
      .map((subArray) => {
        const str = subArray.map((num) => `${target}${num}`).join(':');
        if (isMulti) {
          return `${str},${condition}`;
        }
        return str;
      })
      .join(',');
  }

  /**
   * @description: 상영 시간표 데이터 만들기 ( 주로 이 데이터를 활용해서 다른 시트 데이터 변환 작업 진행 함 )
   *
   * @process:
   * 1. 데이터를 상영시간표 포맷에 맞는 데이터로 변환
   * 1-1. 계열사, 지역명, 극장명, 총좌석수, 상영관명, 포맷 값이 다 일치할 경우 회차 합치는 작업 진행
   * 1-2. 각 계열사별로 배열을 만들고, 각각 계열사에 데이터를 추가
   * 1-3. 그리고 한번 더 데이터 정제, ( 총합 데이터 관련 )
   * 1-3. 그렇게 만들어진 데이터와, 가장 큰 회차 개수를 함께 반환
   */
  createScreenTableData(
    movieTotalList: Array<IMovieCrawlRequest>,
    combineArr: Array<Array<string>>
  ): {
    totalTimeLength: any;
    data: any;
  } {
    let totalTimeLength = 0;
    /*
     *  { header: '지역', key: 'areaNm', columnKey: 'areaNm' },
     *  { header: '극장명', key: 'theaterNm', columnKey: 'theaterNm' },
     *  { header: '비고', key: 'note', columnKey: 'note' },
     *  { header: '관', key: 'hall', columnKey: 'hall' },
     *  { header: '좌석수', key: 'totalSeat', columnKey: 'totalSeat' },
     *  ...timeHeaderList,
     *  { header: '총회차', key: 'totalTimeCnt', columnKey: 'totalTimeCnt' },
     *  { header: '총스크린', key: 'screenCntCnt', columnKey: 'totalScreenCnt' },
     *  { header: '총좌석수', key: 'totalSeatCnt', columnKey: 'totalSeatCnt' },
     * */
    let idx = 0;
    const data = ObjectUtil.newJson<Array<IMovieCrawlRequest>>(movieTotalList).reduce(
      (acc: any, movieTotalInfo: IMovieCrawlRequest) => {
        const { theaterCompany, type, refineAreaNm, theaterNm, hall, totalSeat, time, remainSeat, isSubtitle } =
          movieTotalInfo;

        if (!acc[theaterCompany]) {
          acc[theaterCompany] = [];
        }

        let findIdx = -1; // 같은 홀 정보까지 찾아서, time 정보를 넣기 위해
        let subTitleTxt = '';
        if (combineArr.length > 0) {
          findIdx = acc[theaterCompany].findIndex(
            (d) =>
              d.theaterCompany === theaterCompany &&
              d.areaNm === refineAreaNm &&
              d.theaterNm === theaterNm &&
              d.totalSeat === totalSeat &&
              d.hall === hall
          );
        } else {
          subTitleTxt = isSubtitle ? '자막' : '더빙';

          findIdx = acc[theaterCompany].findIndex(
            (d) =>
              d.theaterCompany === theaterCompany &&
              d.areaNm === refineAreaNm &&
              d.theaterNm === theaterNm &&
              d.totalSeat === totalSeat &&
              d.hall === hall &&
              d.note === subTitleTxt
          );
        }

        if (findIdx !== -1 && time.length > 0) {
          const tg = acc[theaterCompany][findIdx];
          tg[`time_${tg.timeLen + 1}`] = time;
          tg.timeLen = tg.timeLen + 1;
          tg.totalTimeCnt = tg.totalTimeCnt + 1;
          tg.totalSeatCnt = tg.totalSeatCnt + totalSeat;
          tg.totalRemainSeat = tg.totalRemainSeat + (remainSeat === -1 ? totalSeat : remainSeat);
        } else {
          const refineDt: any = {
            theaterCompany,
            areaNm: refineAreaNm,
            theaterNm,
            note: subTitleTxt,
            type,
            totalSeat,
            hall,
            totalTimeCnt: 0,
            screenCntCnt: 1,
            totalSeatCnt: totalSeat,
            timeLen: 0,
            totalRemainSeat: remainSeat === -1 ? totalSeat : remainSeat,
            idx,
          };

          if (time.length > 0) {
            refineDt['time_1'] = time;
            refineDt['totalTimeCnt'] = refineDt['totalTimeCnt'] + 1;
            refineDt['timeLen'] = refineDt['timeLen'] + 1;
          }

          acc[theaterCompany].push(refineDt);
          idx++;
        }

        return acc;
      },
      {}
    );

    function getAllHallToStr(data: any, len: number) {
      const result = Array(len)
        .fill(1)
        .map((d, idx) => data[`time_${idx + 1}`])
        .join(',');

      return result;
    }

    function compareAndCombine(target, sub): { isCombine: boolean; isCompare: boolean } {
      let isCombine = false;
      let isCompare = true;
      const targetTimes = getAllHallToStr(target, target.timeLen).split(',');
      const subTimes = getAllHallToStr(sub, sub.timeLen).split(',');

      subTimes.forEach((subTime, idx) => {
        if (!targetTimes.includes(subTime)) {
          target[`time_${target.timeLen + 1}`] = subTime;
          target['timeLen'] += 1;
          target['totalTimeCnt'] += 1;
          isCombine = true;
          isCompare = false;
        }
      });
      return {
        isCombine,
        isCompare,
      };
    }

    const refineData = Object.keys(data).reduce((acc, theaterCompany) => {
      if (!acc[theaterCompany]) {
        acc[theaterCompany] = [];
      }

      data[theaterCompany].forEach((movieInfo) => {
        /// 계열사 + 상영관 + 상영 시간 개수 동일 시 합치는 부분
        const accIdx = acc[theaterCompany].findIndex(
          (infoInAcc) =>
            getAllHallToStr(infoInAcc, infoInAcc.timeLen) === getAllHallToStr(movieInfo, movieInfo.timeLen) &&
            infoInAcc.theaterNm === movieInfo.theaterNm &&
            infoInAcc.theaterCompany === movieInfo.theaterCompany
        );

        // 판매 좌석수, 좌석 판매율 설정
        movieInfo.soldSeatCnt = movieInfo.totalSeatCnt - movieInfo.totalRemainSeat;
        movieInfo.seatRateAvg = ((movieInfo.soldSeatCnt / movieInfo.totalSeatCnt) * 100).toFixed(1);

        if (accIdx !== -1) {
          const { isCompare, isCombine } = compareAndCombine(acc[theaterCompany][accIdx], movieInfo);

          const isHighAccData = acc[theaterCompany][accIdx].totalSeat > movieInfo.totalSeat;
          acc[theaterCompany][accIdx].hall = isHighAccData ? acc[theaterCompany][accIdx].hall : movieInfo.hall;
          const accMovieInfo = acc[theaterCompany][accIdx];

          accMovieInfo.totalSeat = accMovieInfo.totalSeat += movieInfo.totalSeat;

          accMovieInfo.totalSeatCnt = accMovieInfo.totalSeatCnt += movieInfo.totalSeatCnt;
          accMovieInfo.totalRemainSeat = accMovieInfo.totalRemainSeat += movieInfo.totalRemainSeat;
          accMovieInfo.soldSeatCnt = accMovieInfo.totalSeatCnt - accMovieInfo.totalRemainSeat;
          accMovieInfo.seatRateAvg = ((accMovieInfo.soldSeatCnt / accMovieInfo.totalSeatCnt) * 100).toFixed(1);

          // 판매 좌석수, 좌석 판매율 설정
        } else {
          acc[theaterCompany].push(movieInfo);
        }

        totalTimeLength = totalTimeLength > movieInfo.totalTimeCnt ? totalTimeLength : movieInfo.totalTimeCnt;
      });
      return acc;
    }, {});
    return {
      totalTimeLength,
      data: refineData,
    };
  }

  /**
   * @description: 집계작 및 경쟁작 관련 시트 데이터 생성 함수
   */
  createTotalTableForCompareData(movieName: string, screenTableData: { data: any; totalTimeLength: any }) {
    const totalRemainSeatByArea = {};
    const theaterCntByArea = {};
    const dataKeys = Object.keys(screenTableData.data);
    const testObj = {};
    return dataKeys.reduce((acc, theaterCompany, idx) => {
      screenTableData.data[theaterCompany].forEach((d) => {
        if (!acc[theaterCompany]) {
          acc[theaterCompany] = {
            hallCnt: 0,
            theaterCnt: 0,
            timeCnt: 0,
            screenCnt: 0,
            totalSeatCnt: 0,
            timeAvg: 0,
            seatCntAvg: 0,
            seatRateAvg: 0,
            totalRemainSeat: 0,
          };
        }

        if (!theaterCntByArea[theaterCompany]) {
          theaterCntByArea[theaterCompany] = [];
        }

        if (!totalRemainSeatByArea[theaterCompany]) {
          totalRemainSeatByArea[theaterCompany] = 0;
        }

        if (!theaterCntByArea[theaterCompany].includes(d.theaterNm)) {
          theaterCntByArea[theaterCompany].push(d.theaterNm);
        }

        const hallNm = `${d.theaterNm}_${d.hall}`;
        if (!testObj[theaterCompany]) {
          testObj[theaterCompany] = [];
        }

        testObj[theaterCompany].push(hallNm);

        const { screenCnt, timeCnt, totalSeatCnt, totalRemainSeat } = acc[theaterCompany];

        acc[theaterCompany] = {
          ...acc[theaterCompany],
          timeCnt: timeCnt + d.timeLen,
          screenCnt: screenCnt + d.screenCntCnt,
          totalSeatCnt: totalSeatCnt + d.totalSeatCnt,
          totalRemainSeat: totalRemainSeat + d.totalRemainSeat,
        };

        totalRemainSeatByArea[theaterCompany] += d.totalRemainSeat;
      });
      if (dataKeys.length - 1 === idx) {
        Object.keys(acc).forEach((subTheaterCompany) => {
          const data = acc[subTheaterCompany];

          data.hallCnt = Array.from(new Set(testObj[subTheaterCompany])).length;
          data.theaterCnt = theaterCntByArea[subTheaterCompany].length;
          const { totalSeatCnt, timeCnt, theaterCnt } = data;
          data.seatCntAvg = timeCnt / theaterCnt;
          data.timeAvg = totalSeatCnt / timeCnt;

          const soldSeatCnt = data.totalSeatCnt - data.totalRemainSeat;

          data.seatRateAvg = ((soldSeatCnt / data.totalSeatCnt) * 100).toFixed(1);
        });
      }

      return acc;
    }, {} as { [areaNm: string]: IDataByTheaterCompanyInfo });
  }

  /**
   * @description: 초기 데이터에서 필터링된 데이터만 추출하는 함수
   *
   * @process:
   * 1. 기존 데이터를 가져온다.
   * 2. 지역명이 바뀌지 않았을 가능성을 고려해, 지역명 변환 ( ex] 경기 -> 경강 )
   * 2-1. "자막(2d) => 자막" 으로 변환해주는 작업을 진행한다.
   * 2-2. 극장명이 "CINE de CHEF 센텀" 이 아닌 데이터들만 추출한다.
   */
  testRefineForData(objData: { [screenDt: string]: { [movieNm: string]: Array<IMovieCrawlRequest> } }) {
    const result = {};
    Object.keys(objData).forEach((screenDt) => {
      if (!result[screenDt]) {
        result[screenDt] = {};
      }
      Object.keys(objData[screenDt]).forEach((movieNm) => {
        if (!result[screenDt][movieNm]) {
          result[screenDt][movieNm] = [];
        }
        objData[screenDt][movieNm].forEach((d) => {
          const { theaterCompany, areaNm, theaterNm } = d;
          d.refineAreaNm = getAreaNmFromRefine(theaterNm, theaterCompany as MovieNameEnum) ?? areaNm;
          d.type = d.type.replaceAll(/\([a-zA-Z0-9ㄱ-ㅎ가-힣\s]+\)/g, '');
          if (ObjectUtil.removeSpecialChar(theaterNm) !== ObjectUtil.removeSpecialChar('CINE de CHEF 센텀')) {
            result[screenDt][movieNm].push(d);
          }
        });
      });
    });

    return result;
  }

  /**
   * @description: 엑셀 생성 메인 함수
   *
   * @process:
   * 1. 매개변수를 통해 데이터, 크롤 시간 등을 받아온다.
   * 2. 값이 제대로 들어왔는지 확인한다. ( 없으면 취소 )
   * 3. json 설정 데이터를 토대로 작업을 시작한다.
   * 4. 먼저 각 영화명에 맞는 상영시간표 데이터를 만든다. ( 계열사 별로 배열 형태로 값이 들어가져 있음 )
   * 5. 해당 상영시간표 데이터를 토대로 다른 시트들에서 활용할 데이터를 추가로 만든다.
   * 6. 각각 시트를 만드는 함수를 통해, 각 시트를 생성한다.
   *  6-1. 상영시간표 ( 날짜별 )
   *  6-2. 계열사별
   *  6-3. 지역별
   *  6-4. 집계작 및 경쟁작 멀티3사 비교
   *  6-5. 경쟁작
   *  6-6. 포맷별 요약표
   * 7. 파일을 생성한다.
   */
  async refineDbDataForExcel(
    propertyInfo: IMovieSetting,
    objData: { [screenDt: string]: { [movieNm: string]: Array<IMovieCrawlRequest> } },
    crawlTime: string
  ): Promise<any> {
    objData = this.testRefineForData(objData);

    const { crawlStartDate, crawlEndDate, savePath, movieSettings } = propertyInfo;
    const response: IExcelResponse<any> = {
      status: true,
      statusCd: 200,
      message: '',
      data: null,
    };
    if (movieSettings.length === 0) {
      const message = `[Make Excel]: 영화 정보가 입력되지 않았습니다.`;
      logger.info(message);

      logger.info(message);
      response.message = message;
      response.status = false;
      response.statusCd = 404;

      return response;
    }

    if (!savePath || savePath.length === 0) {
      const message = `[Make Excel]: 엑셀 저장 위치가 입력되지 않았습니다.`;
      logger.info(message);

      logger.info(message);
      response.message = message;

      return response;
    }

    if (!crawlStartDate || !crawlEndDate) {
      const message = `[Make Excel]: 엑셀 시작 날짜 또는 끝나는 날짜가 입력되지 않았습니다.`;
      logger.info(message);

      logger.info(message);
      response.message = message;

      return response;
    }

    function getOriginalMovieNmInObject(object: any, movieNm: string): string {
      const refineMovieNm: string = ObjectUtil.removeSpecialChar(movieNm);
      const ObjKeyArray = Object.keys(object);

      const idx = ObjKeyArray.findIndex((targetMovieNm) => {
        const refineTargetMovieNm = ObjectUtil.removeSpecialChar(targetMovieNm);
        return (
          refineMovieNm === refineTargetMovieNm ||
          (refineMovieNm.length > 0 && refineTargetMovieNm.includes(refineMovieNm))
        );
      });

      return idx === -1 ? movieNm : ObjKeyArray[idx];
    }

    let dataByTheaterCompanyObj = {};
    let dataByAreaObj = {};

    for await (const movieData of movieSettings) {
      const workbook = new ExcelJS.Workbook();
      const compareDataArray: Array<{ screenDt: string; movieArray: Array<{ movieNm: string; data: any }> }> = [];
      const compareDataForTheater: { [screenDt: string]: Array<{ movieNm: string; data: any }> } = {};

      const { rivalMovieNames } = movieData;

      let screenTableDataArray: TScreenTable = [];

      let mainMovieNm = ObjectUtil.removeSpecialChar(movieData.movieName);

      Object.keys(objData).forEach((screenDt) => {
        const screenData = objData[screenDt];
        mainMovieNm = getOriginalMovieNmInObject(screenData, mainMovieNm);
        const mainMovieData = screenData[mainMovieNm] ?? [];
        if (!mainMovieData) {
          console.log(`"${mainMovieData}" 영화명이 잘못되었습니다.`);
          return;
        }

        const screenMainData = this.createScreenTableData(newJson(mainMovieData), []);

        // 경쟁작에 집계작 넣는 부분
        // compareDataForTheater[screenDt].push({
        //   movieNm: mainMovieNm,
        //   data: this.createScreenTableData(newJson(mainMovieData), []),
        // });

        screenTableDataArray.push({ screenDt, data: screenMainData });
        const test = this.createTotalTableForCompareData(
          mainMovieNm,
          this.createScreenTableData(newJson(mainMovieData), [])
        );
        const idx = compareDataArray.findIndex((d) => d.screenDt === screenDt);

        if (idx === -1) {
          compareDataArray.push({ screenDt: screenDt, movieArray: [{ movieNm: mainMovieNm, data: test }] });
        } else {
          compareDataArray[idx].movieArray.push({ movieNm: mainMovieNm, data: test });
        }

        rivalMovieNames.forEach((rivalMovieNm) => {
          if (!compareDataForTheater[screenDt]) {
            compareDataForTheater[screenDt] = [];
          }

          let refineRivalMovieNm = ObjectUtil.removeSpecialChar(rivalMovieNm);
          refineRivalMovieNm = getOriginalMovieNmInObject(screenData, refineRivalMovieNm);

          const rivalMovieData = screenData[refineRivalMovieNm] ?? [];
          if (!rivalMovieData) {
            console.log(`"${mainMovieData}" 영화명이 잘못되었습니다.`);
            return;
          }

          const screenRivalData = this.createScreenTableData(rivalMovieData, []);
          const rivalCompareData = this.createTotalTableForCompareData(refineRivalMovieNm, screenRivalData);
          const rivalIdx = compareDataArray.findIndex((d) => d.screenDt === screenDt);

          if (rivalIdx === -1) {
            compareDataArray.push({
              screenDt: screenDt,
              movieArray: [{ movieNm: refineRivalMovieNm, data: rivalCompareData }],
            });
          } else {
            compareDataArray[rivalIdx].movieArray.push({ movieNm: refineRivalMovieNm, data: rivalCompareData });
          }

          compareDataForTheater[screenDt].push({
            movieNm: rivalMovieNm,
            data: screenRivalData,
          });
        });

        dataByTheaterCompanyObj = {
          ...dataByTheaterCompanyObj,
          ...this.createDataByTheaterCompany(mainMovieData, screenDt),
        };

        dataByAreaObj = {
          ...dataByAreaObj,
          ...this.createDataByArea(mainMovieData, screenDt),
        };
      });

      for await (const { screenDt, data } of screenTableDataArray) {
        logger.info(`[Excel] 상영시간표(${mainMovieNm} / ${screenDt}) 제작중..`);
        await this.createScreenTableSheet(workbook, mainMovieNm, screenDt, data, crawlTime);
      }

      logger.info(`[Excel] "계열사별" 시트(${mainMovieNm}) 제작중..`);
      await this.createDataByTheaterCompanySheet(workbook, dataByTheaterCompanyObj);

      logger.info(`[Excel] "지역별" 시트(${mainMovieNm}) 제작중..`);
      await this.createDataByAreaSheet(workbook, dataByAreaObj);

      logger.info(
        `[Excel] "집계작 및 경쟁작 멀티3사 비교" 시트(${mainMovieNm} / ${rivalMovieNames.join(',')}) 제작중..`
      );
      await this.createDataByCompareSheet(workbook, compareDataArray);

      logger.info(`[Excel] "경쟁작" 시트(${mainMovieNm} / ${rivalMovieNames.join(',')}) 제작중..`);
      await this.createDataByCompareForTheaterCompany(workbook, compareDataForTheater, crawlTime);

      logger.info(`[Excel] "포맷별 요약표" 시트(${mainMovieNm} / ${rivalMovieNames.join(',')}) 제작중..`);
      await this.createDataByFormat(workbook, screenTableDataArray, crawlTime);

      const filename = `excel_${mainMovieNm}__${crawlStartDate}-${crawlEndDate}.xlsx`;
      try {
        await workbook.xlsx.writeFile(`${savePath}${path.sep}${filename}`);
      } catch (err) {
        logger.error(`[Make Excel]: 엑셀로 만드는 도중 에러가 발생. filename: '${filename}'`);
        console.log('err', err);
      }
      console.log('success');
      logger.info(`Excel finish... ${filename}`);
    }
    logger.info('Excel finish... all');
  }

  /**
   * @description: 엑셀에 값을 넣을때 비어이있는 공간에 빈 문자열을 넣기 위해 만든 함수이지만,
   *               count 샐 때 이 공백도 포함되기 떄문에 내부 로직은 주석 처리 해둔 상태.
   */
  addEmptyValue(workSheet: ExcelJS.Worksheet, row: ExcelJS.Row) {
    workSheet.columns.forEach((column) => {
      const cell = row.getCell(column.key);
      if (cell.value === null) {
        // cell.value = '';
      }
    });
  }

  /**
   * @description: 숫자 관련 데이터에 천 단위로 ,(콤마)를 붙여주는 함수
   */
  addCommaAndFixed(info: any, targetArray: Array<any>) {
    const dataForAdd = JSON.parse(JSON.stringify(info));

    targetArray.forEach((key) => {
      if (key.includes('timeLen')) {
        dataForAdd[key] = dataForAdd[key] ? dataForAdd[key].toLocaleString() : '';
      }
      if (key.includes('totalSeat')) {
        dataForAdd[key] = dataForAdd[key] ? dataForAdd[key].toLocaleString() : '';
      }
      if (key.includes('totalSeatCnt')) {
        dataForAdd[key] = dataForAdd[key] ? dataForAdd[key].toLocaleString() : '';
      }
      if (key.includes('soldSeatCnt')) {
        dataForAdd[key] = dataForAdd[key] ? dataForAdd[key].toLocaleString() : '';
      }
      if (key.includes('seatRateAvg')) {
        const seatRateAvg = dataForAdd[key] ? Number(dataForAdd[key]).toFixed(1) : '';

        dataForAdd[key] = isNaN(Number(seatRateAvg)) ? '0' : seatRateAvg;
      }
    });
    return dataForAdd;
  }

  /**
   * @description: 경쟁작 시트를 만드는 함수
   *
   * @process:
   */
  async createDataByCompareForTheaterCompany(
    workbook: ExcelJS.Workbook,
    compareData: { [screenDt: string]: Array<{ movieNm: string; data: any }> },
    crawlTime: string
  ) {
    const workSheet: ExcelJS.Worksheet = setWorkSheet(workbook, '경쟁작', {
      views: [
        {
          state: 'frozen',
          ySplit: 4,
        },
      ],
    });

    if (Object.keys(compareData).length === 0) {
      logger.info(`[Excel -> 경쟁작] 경쟁작 영화명이 설정되어 있지 않아 빈 시트만 생성됩니다.`);
      return;
    }

    const maxObj = Object.keys(compareData).reduce(
      (acc, screenDt, idx) => {
        const value = compareData[screenDt];

        const len = value.length;
        if (acc.cnt < len) {
          acc = {
            cnt: len,
            screenDt,
          };
        }
        return acc;
      },
      { cnt: 0, screenDt: '' }
    );

    console.log(compareData, 'compareData');
    console.log(maxObj, 'maxObj');

    const movieNames = compareData[maxObj.screenDt].map((d) => d.movieNm);
    const header = [
      { header: '상영일자', key: 'screenDt', columnKey: 'screenDt', width: 19 },
      { header: '영화관', key: 'theaterNm', columnKey: 'theaterNm' },
      ...movieNames.reduce((acc, movieNm) => {
        acc.push({ header: '상영관', key: `${movieNm}_hall`, columnKey: `${movieNm}_hall` });
        acc.push({ header: '회차', key: `${movieNm}_timeLen`, columnKey: `${movieNm}_timeLen` });
        acc.push({ header: '좌석수', key: `${movieNm}_totalSeat`, columnKey: `${movieNm}_totalSeat` });
        acc.push({ header: '총좌석수', key: `${movieNm}_totalSeatCnt`, columnKey: `${movieNm}_totalSeatCnt` });
        acc.push({ header: '판매좌석수', key: `${movieNm}_soldSeatCnt`, columnKey: `${movieNm}_soldSeatCnt` });
        acc.push({
          header: '판매좌석률',
          key: `${movieNm}_seatRateAvg`,
          columnKey: `${movieNm}_seatRateAvg`,
          width: 15,
        });
        return acc;
      }, []),
    ];

    const defaultColumnCnt = 2;
    const dataHeaderCnt = movieNames.length;

    const setTopHeadersOption = {
      defaultColumn: {
        background: 'd9e1f2',
        valueLen: defaultColumnCnt,
        width: 15,
        values: ['상영일자', '예매오픈현황'],
      },
      defaultHeaders: {
        background: 'd9e1f2',
        valueLen: dataHeaderCnt,
        repeatCnt: 6,
        values: movieNames ?? [],
      },
    };

    workSheet.columns = header;
    // 기존에는 header로 설정된 row는 1번째 줄로 고정된다.
    // 특정 원하는 값을 1번째 줄에 놓고, header의 설정된 row을 아래로 바꾸게 해주기 위해 아래와 같이 코드 작성하였음.
    workSheet.spliceRows(
      4,
      2,
      header.map((d) => d.header)
    );
    // 상단의 2개의 헤더를 설정해야 하기에 관련 로직을 함수로 분리
    this.setTopHeaders(workSheet, setTopHeadersOption, 3);
    workSheet.getColumn('A').width = 15;
    workSheet.getColumn('B').width = 23;
    workSheet.getRow(1).eachCell((cell, num) => {
      cell.value = '';
    });
    const crawlTimeCell = workSheet.getCell('A:1');
    crawlTimeCell.value = crawlTime;
    workSheet.mergeCells(`A1:C1`);
    // 커스텀으로 만든 함수로, 폰트 굵기, 사이즈, 컬러등을 조정 가능하다.
    setFont2(crawlTimeCell, { bold: true, size: 12, color: 'ff3a00' });

    const namesInfoList = [];

    const refineData = Object.keys(compareData).reduce((acc, screenDt) => {
      if (!acc[screenDt]) {
        acc[screenDt] = {
          cgv: [],
          lotte: [],
          mega: [],
        };
      }

      const screenDtValue = compareData[screenDt];

      screenDtValue.forEach(({ movieNm, data: { data } }) => {
        Object.keys(data).forEach((theaterCompany) => {
          const theaterValue = data[theaterCompany];
          theaterValue.forEach((info) => {
            const {
              screenCntCnt,
              timeLen,
              totalSeat,
              totalSeatCnt,
              hall,
              soldSeatCnt,
              seatRateAvg,
              theaterNm,
              note,
              type,
            } = info;

            if (!acc[screenDt][theaterCompany]) {
              acc[screenDt][theaterCompany] = [];
            }

            const searchHall = theaterNm + hall;

            const accTheaterIdx = acc[screenDt][theaterCompany].findIndex((data) => data.searchHall === searchHall);

            if (theaterCompany === 'cgv' && theaterNm.includes('강릉') && screenDt === '2023-05-30') {
              console.log('===============================');
              console.log('info', info, screenDt);
              console.log(acc[screenDt][theaterCompany][accTheaterIdx], 'combine target info', screenDt, 'before');
            }

            const refineInfo = {
              [`${movieNm}_hall`]: hall.length === 0 ? null : hall,
              [`${movieNm}_searchHall`]: searchHall,
              [`${movieNm}_timeLen`]: timeLen,
              [`${movieNm}_totalSeat`]: totalSeat,
              [`${movieNm}_totalSeatCnt`]: totalSeatCnt,
              [`${movieNm}_soldSeatCnt`]: soldSeatCnt,
              [`${movieNm}_seatRateAvg`]: seatRateAvg,
              [`${movieNm}_note`]: note,
            };

            Object.keys(refineInfo).forEach((key) => {
              if (!namesInfoList.includes(key)) {
                namesInfoList.push(key);
              }

              if (accTheaterIdx !== -1) {
                const originalData = acc[screenDt][theaterCompany][accTheaterIdx];
                if (originalData[key] && refineInfo[key]) {
                  if (
                    key.includes('_timeLen') ||
                    key.includes('_totalSeat') ||
                    key.includes('_totalSeatCnt') ||
                    key.includes('_soldSeatCnt')
                  ) {
                    refineInfo[key] = refineInfo[key] + originalData[key];
                  }
                }
              }
            });

            refineInfo[`${movieNm}_seatRateAvg`] =
              (refineInfo[`${movieNm}_soldSeatCnt`] / refineInfo[`${movieNm}_totalSeatCnt`]) * 100;

            if (theaterCompany === 'cgv' && theaterNm.includes('강릉') && screenDt === '2023-05-30') {
              console.log(refineInfo, 'refineInfo', 'after');
            }

            if (accTheaterIdx === -1) {
              acc[screenDt][theaterCompany].push({
                movieNm,
                screenDt,
                theaterCompany,
                theaterNm,
                hall,
                searchHall,
                ...refineInfo,
              });
            } else {
              acc[screenDt][theaterCompany][accTheaterIdx] = {
                ...acc[screenDt][theaterCompany][accTheaterIdx],
                ...refineInfo,
              };
            }
          });
        });
      });

      return acc;
    }, {} as { [screenDt: string]: { cgv: Array<any>; mega: Array<any>; lotte: Array<any> } });

    let addIdx = 3;
    let dataLength = 0;
    let firstIdx = 5;
    Object.keys(refineData).forEach((screenDt) => {
      const screenDtValue = refineData[screenDt];
      const totalSumObj = {};
      const totalTheaterNmList = [];
      Object.keys(screenDtValue).forEach((theaterCompany) => {
        const array = screenDtValue[theaterCompany];
        const sumObj = {};
        const theaterNmList = [];
        array
          .sort((d1, d2) => (d1.theaterNm < d2.theaterNm ? -1 : 1))
          .forEach((info) => {
            namesInfoList.forEach((key) => {
              const [movieNm, type] = key.split('_');
              if (key.includes('hall')) {
                const hallNm = movieNm + '_' + info.theaterNm + '_' + info.hall;
                if (!sumObj[key]) {
                  sumObj[key] = 0;
                }

                if (info[key]) {
                  sumObj[key] += 1;
                  theaterNmList.push(hallNm);
                }

                if (!totalTheaterNmList.includes(hallNm)) {
                  totalTheaterNmList.push(hallNm);
                }
              } else if (key.includes('seatRateAvg')) {
                const sumObjTotalRateAvg =
                  ((sumObj[`${movieNm}_soldSeatCnt`] ?? 0) / (sumObj[`${movieNm}_totalSeatCnt`] ?? 0)) * 100;

                sumObj[key] = sumObjTotalRateAvg === 0 ? 0 : sumObjTotalRateAvg.toFixed(1);

                const totalSoldSeatCnt =
                  ((totalSumObj[`${movieNm}_soldSeatCnt`] ?? 0) / (totalSumObj[`${movieNm}_totalSeatCnt`] ?? 0)) * 100;
                totalSumObj[key] = totalSoldSeatCnt === 0 ? '0' : totalSoldSeatCnt.toFixed(1);
              } else if (key !== 'note') {
                if (!sumObj[key]) {
                  sumObj[key] = 0;
                }
                sumObj[key] += info[key] ?? 0;

                if (!totalSumObj[key]) {
                  totalSumObj[key] = 0;
                }
                totalSumObj[key] += info[key] ?? 0;
              }
            });

            Object.keys(info).forEach((key) => {
              if (key.includes('hall')) {
                info[key] = info[key].trim();
              }
            });

            const row = workSheet.addRow({
              ...this.addCommaAndFixed(info, namesInfoList),
              theaterNm: this.refineTheaterNm(theaterCompany, info.theaterNm),
            });
            this.addEmptyValue(workSheet, row);

            addIdx++;
          });
        dataLength += array.length + 1;
        addIdx += 1;

        const row = workSheet.addRow({
          screenDt,
          ...this.addCommaAndFixed(sumObj, namesInfoList),
          theaterNm: `${theaterCompany.toUpperCase()} 계`,
        });

        namesInfoList.forEach((key) => {
          if (key.includes('_hall')) {
            if (!totalSumObj[key]) {
              totalSumObj[key] = 0;
            }
            totalSumObj[key] += sumObj[key];
          }
        });
        firstIdx = row.number + 1;

        addIdx += 1;
      });
      dataLength += 1;
      addIdx += 1;

      const row = workSheet.addRow({
        screenDt,
        theaterNm: `___`,
        ...this.addCommaAndFixed(totalSumObj, namesInfoList),
      });

      firstIdx = row.number + 1;

      addIdx += 1;
    });

    const headerTopRow = workSheet.findRow(3);
    const headerRow = workSheet.findRow(4);
    header.forEach(({ key }) => {
      const headerTopCell = headerTopRow.getCell(key);
      const cell = headerRow.getCell(key);

      if (key.includes('timeLen')) {
        setBorder2(headerTopCell, { left: 'medium', right: 'medium' });
      } else {
        setBorder2(headerTopCell, { left: 'thin', right: 'thin' });
      }

      setBorder2(headerTopCell, { bottom: 'medium', top: 'thin' });

      setBorder(workSheet, cell, 'thin');
      setFill(cell, 'd9e1f2');
      setCenter(cell);
      setFont2(cell, { bold: true });

      if (key.includes('hall')) {
        setBorder2(cell, { left: 'medium' });
        workSheet.getColumnKey(key).width = 17;
      } else if (key.includes('seatRateAvg')) {
        setBorder2(cell, { right: 'medium' });
        workSheet.getColumnKey(key).width = 10;
      } else if (key.includes('soldSeatCnt')) {
        workSheet.getColumnKey(key).width = 10;
      }
    });
    workSheet.findRows(5, workSheet.lastRow.number).forEach((row) => {
      const checkCell = row.getCell('theaterNm');
      if (
        /^([a-zA-Z0-9ㄱ-ㅎ가-힣]+\s계)$/g.test(String(checkCell?.value ?? '')) ||
        String(checkCell?.value ?? '') === '___'
      ) {
        const isTotalCell = checkCell?.value === '___';
        header.forEach((data) => {
          const { columnKey } = data;
          const cell = row.getCell(data.columnKey);

          setBorder(workSheet, cell, 'thin');
          setFill(cell, isTotalCell ? 'ffc000' : '92cddc');
          setFont2(cell, { size: 10, bold: true, color: '000000' });
          setCenter(cell);

          if (columnKey.includes('hall')) {
            setBorder2(cell, { left: 'medium' });
          } else if (columnKey.includes('seatRateAvg')) {
            setBorder2(cell, { right: 'medium' });
          }
        });
      } else {
        header.forEach((data) => {
          const { columnKey } = data;
          const cell = row.getCell(data.columnKey);

          setBorder(workSheet, cell, 'thin');
          setFont2(cell, { size: 10, color: '000000' });
          setCenter(cell);

          if (columnKey === 'screenDt') {
            setFill(cell, 'f7f9e1');
            setFont2(cell, { color: '203764', bold: true });
          } else if (columnKey === 'theaterNm') {
            setFill(cell, 'ebf6a6');
            setFont2(cell, { color: '203764', bold: true });
            setBorder2(cell, { right: 'medium' });
          } else if (columnKey.includes('hall')) {
            setBorder2(cell, { left: 'medium' });
          } else if (columnKey.includes('seatRateAvg')) {
            setBorder2(cell, { right: 'medium' });
          }
        });
      }
    });
  }

  /**
   * @description: 집계작 및 경쟁작 멀티3사 비교 시트를 만드는 함수
   * - 헤더를 설정할 때: 상영일자 / 영화관 / 영화이름1_상영관 / 영화이름1_회차 ... / 영화이름2_상영관 / 영화이름2_회차 ...
   *   위의 형식으로 헤더를 구성하였고, 위의 키값에 맞게 데이터를 조정 후 삽입하는 형식으로 진행.
   *
   * @process:
   */
  async createDataByCompareSheet(
    workbook: ExcelJS.Workbook,
    compareDataArray: Array<{ screenDt: string; movieArray: Array<{ movieNm: string; data: any }> }>
  ) {
    const maxObj = compareDataArray.reduce(
      (acc, data, idx) => {
        const len = data.movieArray.length;
        if (acc.cnt < len) {
          acc = {
            cnt: len,
            idx,
          };
        }
        return acc;
      },
      { cnt: 0, idx: 0 }
    );
    const movieNames = compareDataArray[maxObj.idx].movieArray.map((d) => d.movieNm);
    const workSheet: ExcelJS.Worksheet = setWorkSheet(workbook, '집계작 및 경쟁작 멀티3사 비교');
    const header = [
      { header: '상영일자', key: 'screenDt', columnKey: 'screenDt', width: 19 },
      { header: '영화관', key: 'theaterCompany', columnKey: 'theaterCompany' },
      ...movieNames.reduce((acc, movieNm) => {
        acc.push({ header: '상영관', key: `${movieNm}_hallCnt`, columnKey: `${movieNm}_hallCnt` });
        acc.push({ header: '회차', key: `${movieNm}_timeCnt`, columnKey: `${movieNm}_timeCnt` });
        acc.push({ header: '총좌석수', key: `${movieNm}_totalSeatCnt`, columnKey: `${movieNm}_totalSeatCnt` });
        acc.push({
          header: '평균 좌판율',
          key: `${movieNm}_seatRateAvg`,
          columnKey: `${movieNm}_seatRateAvg`,
          width: 15,
        });
        return acc;
      }, []),
    ];

    const defaultColumnCnt = 2;
    const dataHeaderCnt = movieNames.length;
    const topHeaders = ['상영일자', '예매오픈현황', ...movieNames];

    const setTopHeadersOption = {
      defaultColumn: {
        background: 'd9e1f2',
        valueLen: defaultColumnCnt,
        width: 15,
        values: ['상영일자', '예매오픈현황'],
      },
      defaultHeaders: {
        background: 'd9e1f2',
        valueLen: dataHeaderCnt,
        repeatCnt: 4,
        values: movieNames ?? [],
      },
    };

    const refineData = compareDataArray.reduce((acc, { screenDt, movieArray }) => {
      const result = {
        [screenDt]: [],
      };

      const movieNames = [];
      movieArray.forEach(({ movieNm, data }) => {
        if (!movieNames.includes(movieNm)) {
          movieNames.push(movieNm);
        }

        this.sortObjKeyAsc(data).forEach((theaterCompany) => {
          const {
            theaterCnt,
            timeCnt,
            screenCnt,
            totalSeatCnt,
            timeAvg,
            seatCntAvg,
            seatRateAvg,
            totalRemainSeat,
            hallCnt,
          } = data[theaterCompany];

          theaterCompany = `${theaterCompany.toUpperCase()} 계`;

          const dd = {
            screenDt,
            theaterCompany,
            [`${movieNm}_theaterCnt`]: theaterCnt,
            [`${movieNm}_hallCnt`]: hallCnt,
            [`${movieNm}_timeCnt`]: timeCnt,
            [`${movieNm}_screenCnt`]: screenCnt,
            [`${movieNm}_totalSeatCnt`]: totalSeatCnt,
            [`${movieNm}_timeAvg`]: timeAvg,
            [`${movieNm}_seatCntAvg`]: seatCntAvg,
            [`${movieNm}_seatRateAvg`]: seatRateAvg,
            [`${movieNm}_totalRemainSeat`]: totalRemainSeat,
          };

          const theaterCompanyIdx = result[screenDt].findIndex((d) => {
            return d.theaterCompany === theaterCompany;
          });

          if (theaterCompanyIdx === -1) {
            result[screenDt].push({
              theaterCompany: theaterCompany,
              data: dd,
            });
          } else {
            const findMovies = Object.keys(result[screenDt][theaterCompanyIdx].data).map((d) => d.split('_')[0] ?? '');
            if (findMovies.filter((d) => d === movieNm).length > 0) {
              const ddd = result[screenDt][theaterCompanyIdx].data;
              ddd[`${movieNm}_timeCnt`] = (ddd[`${movieNm}_timeCnt`] ?? 0) + dd[`${movieNm}_timeCnt`];
              ddd[`${movieNm}_totalSeatCnt`] = (ddd[`${movieNm}_totalSeatCnt`] ?? 0) + dd[`${movieNm}_totalSeatCnt`];
              ddd[`${movieNm}_totalRemainSeat`] =
                (ddd[`${movieNm}_totalRemainSeat`] ?? 0) + dd[`${movieNm}_totalRemainSeat`];
              const soldSeatCnt =
                ddd[`${movieNm}_totalSeatCnt`] -
                (ddd[`${movieNm}_totalRemainSeat`] === 0
                  ? ddd[`${movieNm}_totalSeatCnt`]
                  : ddd[`${movieNm}_totalRemainSeat`]);
              ddd[`${movieNm}_seatRateAvg`] = ((soldSeatCnt / ddd[`${movieNm}_totalSeatCnt`]) * 100).toFixed(1);
            } else {
              result[screenDt][theaterCompanyIdx].data = {
                ...result[screenDt][theaterCompanyIdx].data,
                ...dd,
              };
            }
          }
        });
      });

      Object.keys(result)
        .sort((a1, a2) => (a1 < a2 ? -1 : 1))
        .forEach((screenDt) => {
          const dataArray = result[screenDt];
          dataArray.forEach((d) => {
            acc.push(d.data);
          });
        });

      const totalObj = {};
      const array = result[screenDt];
      array.forEach(({ data }) => {
        const movieNames = Array.from(
          new Set(
            Object.keys(data)
              .filter((key) => !['theaterCompany', 'screenDt'].includes(key))
              .map((key) => key.split('_')[0])
          )
        );

        movieNames.forEach((movieNm) => {
          totalObj[`${movieNm}_hallCnt`] = (totalObj[`${movieNm}_hallCnt`] ?? 0) + data[`${movieNm}_hallCnt`];
          totalObj[`${movieNm}_totalRemainSeat`] =
            (totalObj[`${movieNm}_totalRemainSeat`] ?? 0) + data[`${movieNm}_totalRemainSeat`];
          totalObj[`${movieNm}_theaterCnt`] = (totalObj[`${movieNm}_theaterCnt`] ?? 0) + data[`${movieNm}_theaterCnt`];
          totalObj[`${movieNm}_timeCnt`] = (totalObj[`${movieNm}_timeCnt`] ?? 0) + data[`${movieNm}_timeCnt`];
          totalObj[`${movieNm}_totalSeatCnt`] =
            (totalObj[`${movieNm}_totalSeatCnt`] ?? 0) + data[`${movieNm}_totalSeatCnt`];
        });
      });

      movieNames.forEach((movieNm) => {
        const totalSeatCnt = totalObj[`${movieNm}_totalSeatCnt`] ?? 0;
        const totalRemainSeat = totalObj[`${movieNm}_totalRemainSeat`] ?? 0;
        const theaterCnt = totalObj[`${movieNm}_theaterCnt`] ?? 0; // 상영관
        const timeCnt = totalObj[`${movieNm}_timeCnt`] ?? 0; // 회차
        const soldSeatCnt = totalSeatCnt - (totalRemainSeat === 0 ? totalSeatCnt : totalRemainSeat); // 남은 좌석수
        const seatRateAvg = (totalObj[`${movieNm}_seatRateAvg`] = (soldSeatCnt / totalSeatCnt) * 100); // 평균 좌판율
        const hallCnt = totalObj[`${movieNm}_hallCnt`] ?? 0;

        totalObj[`${movieNm}_hallCnt`] = hallCnt.toLocaleString() + ' 개관';
        totalObj[`${movieNm}_seatRateAvg`] = (seatRateAvg ? seatRateAvg.toFixed(1) : '0') + '%';
        totalObj[`${movieNm}_theaterCnt`] = theaterCnt.toLocaleString() + ' 개관';
        totalObj[`${movieNm}_timeCnt`] = timeCnt.toLocaleString() + ' 회';
        totalObj[`${movieNm}_totalSeatCnt`] = totalSeatCnt.toLocaleString() + ' 석';
      });

      acc.push({
        total: true,
        ...totalObj,
        screenDt,
        theaterCompany: '',
      });

      return acc;
    }, []);

    workSheet.columns = header;
    // 기존에는 header로 설정된 row는 1번째 줄로 고정된다.
    // 특정 원하는 값을 1번째 줄에 놓고, header의 설정된 row을 아래로 바꾸게 해주기 위해 아래와 같이 코드 작성하였음.
    workSheet.spliceRows(
      2,
      0,
      header.map((d) => d.header)
    );
    this.setTopHeaders(workSheet, setTopHeadersOption);
    workSheet.getColumn(1).width = 17;
    workSheet.getColumn(2).width = 17;

    const startRow = 3;
    let checkScreenDt = refineData[0].screenDt;
    let checkScreenDtIdx = startRow;

    refineData.forEach((data, idx) => {
      const dataForAdd = JSON.parse(JSON.stringify(data));
      movieNames.forEach((movieNm) => {
        dataForAdd[`${movieNm}_hallCnt`] = dataForAdd[`${movieNm}_hallCnt`]?.toLocaleString() || '';
        dataForAdd[`${movieNm}_theaterCnt`] = dataForAdd[`${movieNm}_theaterCnt`]?.toLocaleString() || '';
        dataForAdd[`${movieNm}_timeCnt`] = dataForAdd[`${movieNm}_timeCnt`]?.toLocaleString() || '';
        dataForAdd[`${movieNm}_totalSeatCnt`] = dataForAdd[`${movieNm}_totalSeatCnt`]?.toLocaleString() || '';
        if (dataForAdd[`${movieNm}_seatRateAvg`] && !dataForAdd[`${movieNm}_seatRateAvg`].includes('%')) {
          dataForAdd[`${movieNm}_seatRateAvg`] = (dataForAdd[`${movieNm}_seatRateAvg`] ?? 0) + '%';
        }
      });
      const addedRow = workSheet.addRow(dataForAdd);
      this.addEmptyValue(workSheet, addedRow);

      if (data.screenDt !== checkScreenDt) {
        checkScreenDt = data.screenDt;
        workSheet.mergeCells(`A${checkScreenDtIdx}:A${idx + startRow - 1}`);
        checkScreenDtIdx = idx + startRow;
      }

      if (refineData.length - 1 === idx) {
        workSheet.mergeCells(`A${checkScreenDtIdx}:A${idx + startRow}`);
      }

      if (data?.total) {
        addedRow.eachCell((cell, num) => {
          if (num > 1) {
            setFill(cell, 'ffe699');
          }
        });
      }
    });

    // 상단 1~2번째 줄인 헤더 부분에 스타일을 입히기
    workSheet.findRows(1, 2).forEach((row, idx) => {
      header.forEach(({ key }) => {
        const cell = row.getCell(key);
        setFill(cell, 'd9e1f2');
        setBorder(workSheet, cell, 'thin');
        setCenter(cell);

        if (key.includes('screenCnt')) {
          setBorder2(cell, { left: 'medium' });
        } else if (key.includes('seatRateAvg')) {
          setBorder2(cell, { right: 'medium' });
        }
        setFont2(cell, { bold: true });
      });
      if (idx === 0) {
        // 집계작의 메인 영화의 색 변경
        const cell = row.getCell('C');
        setFill(cell, 'ffe699');
      }
    });
    // 데이터가 들어가는 부분의 스타일 조정
    // workSheet.lastRow.number = exceljs에서 제공되는 변수로, addRow 등을 통해 최종적으로 추가된 row를 가져오고, 해당 row의 number을 가져올 수 있다.
    workSheet.findRows(3, workSheet.lastRow.number).forEach((row) => {
      const checkCell = row.getCell('theaterCompany');
      if (String(checkCell?.value ?? '') === '') {
        row.eachCell((cell, num) => {
          setFont2(cell, { bold: true });
        });
      }

      header.forEach(({ key }) => {
        const cell = row.getCell(key);
        setBorder(workSheet, cell, 'thin');
        setCenter(cell);

        if (key === 'theaterCompany') {
          setFont2(cell, { bold: true });
        }

        if (key.includes('screenCnt')) {
          setBorder2(cell, { left: 'medium' });
        } else if (key.includes('seatRateAvg')) {
          setBorder2(cell, { right: 'medium' });
        } else if (key === 'screenDt') {
          setFill(cell, 'd9e1f2');
          setCenter(cell);
          setBold(cell);
        }
      });
    });
  }

  setTopHeaders(
    workSheet: ExcelJS.Worksheet,
    setTopHeadersOption: {
      defaultColumn: {
        background: string;
        valueLen: string | number;
        width: number;
        values: Array<string>;
      };
      defaultHeaders: {
        background: string;
        valueLen: string | number;
        repeatCnt: number | string;
        values: Array<string>;
      };
    },
    row = 1
  ) {
    const { defaultHeaders, defaultColumn } = setTopHeadersOption;
    const repeatCnt = Number(defaultHeaders.valueLen) + Number(defaultColumn.valueLen);

    let columnIdx = 1;
    for (let i = 1; i <= Number(defaultColumn.valueLen); i++) {
      const cell: ExcelJS.Cell = workSheet.getCell(`${numberToColumn(i)}${row}`);
      setCenter(cell);
      setFill(cell, defaultColumn.background);
      setFont2(cell, { size: 14, bold: true });
      cell.value = defaultColumn.values[i - 1];
    }

    columnIdx += Number(defaultColumn.valueLen);
    let len = Number(defaultHeaders.valueLen) + columnIdx;
    for (let i = columnIdx, forValueIdx = 0; i < len; i++) {
      const repeatCnt = Number(defaultHeaders.repeatCnt);
      const cell: ExcelJS.Cell = workSheet.getCell(`${numberToColumn(i)}${row}`);
      // TODO - 값 추가 안됨
      workSheet.mergeCells(`${numberToColumn(i)}${row}:${numberToColumn(i + repeatCnt - 1)}${row}}`);
      setCenter(cell);
      setFill(cell, defaultHeaders.background);
      setFont2(cell, { size: 14, bold: true, color: '202124' });
      cell.value = defaultHeaders.values[forValueIdx];

      forValueIdx++;
      i += repeatCnt - 1;
      len += repeatCnt - 1;
    }
  }

  async createDataByAreaSheet(workbook: ExcelJS.Workbook, dataByTheaterCompanyObj: any) {
    const workSheet: ExcelJS.Worksheet = setWorkSheet(workbook, '지역별');
    const header = [
      { header: '상영일자', key: 'screenDt', columnKey: 'screenDt', width: 19 },
      { header: '지역', key: 'refineAreaNm', columnKey: 'refineAreaNm' },
      { header: '극장수', key: 'theaterCnt', columnKey: 'theaterCnt' },
      { header: '상영회차', key: 'timeCnt', columnKey: 'timeCnt' },
      { header: '상영관수', key: 'hallCnt', columnKey: 'hallCnt' },
      { header: '총 좌석수', key: 'totalSeatCnt', columnKey: 'totalSeatCnt' },
      { header: '평균회차', key: 'timeAvg', columnKey: 'timeAvg' },
      { header: '평균좌석수', key: 'seatCntAvg', columnKey: 'seatCntAvg', width: 15 },
      { header: '평균 좌판율', key: 'seatRateAvg', columnKey: 'seatRateAvg', width: 15 },
    ];

    workSheet.columns = header;
    let addIdx = 0;
    let dataLength = 0;

    for await (const screenDt of Object.keys(dataByTheaterCompanyObj)) {
      let smallDataLength = 0;
      const screenData = dataByTheaterCompanyObj[screenDt];
      let sum_theaterCnt = 0;
      let sum_timeCnt = 0;
      let sum_hallCnt = 0;
      let sum_totalSeatCnt = 0;
      let sum_timeAvg = 0;
      let sum_remainSeatCnt = 0;

      for await (const refineAreaNm of this.sortObjKeyAsc(screenData)) {
        const value = { ...screenData[refineAreaNm], screenDt, refineAreaNm };
        sum_theaterCnt += value.theaterCnt;
        sum_timeCnt += value.timeCnt;
        sum_hallCnt += value.hallCnt;
        sum_totalSeatCnt += value.totalSeatCnt;
        sum_remainSeatCnt += value.totalRemainSeat;
        sum_timeAvg += value.timeAvg;

        const soldSeatCnt = value.totalSeatCnt - value.totalRemainSeat;
        value.seatRateAvg = ((soldSeatCnt / value.totalSeatCnt) * 100).toFixed(1);
        workSheet.addRow({
          ...value,
          theaterCnt: value.theaterCnt.toLocaleString() ?? 0,
          timeCnt: value.timeCnt.toLocaleString() ?? 0,
          hallCnt: value.hallCnt.toLocaleString() ?? 0,
          totalSeatCnt: value.totalSeatCnt.toLocaleString() ?? 0,
          timeAvg: value.timeAvg.toFixed(1),
          seatCntAvg: value.seatCntAvg.toFixed(1),
        });
        addIdx++;
        dataLength++;
        smallDataLength++;
      }
      addIdx += 1;

      const totalSoldSeatCnt = sum_totalSeatCnt - sum_remainSeatCnt;

      const row = workSheet.addRow({
        screenDt,
        refineAreaNm: '계',
        theaterCnt: sum_theaterCnt.toLocaleString() ?? 0,
        timeCnt: sum_timeCnt.toLocaleString() ?? 0,
        hallCnt: sum_hallCnt.toLocaleString() ?? 0,
        totalSeatCnt: sum_totalSeatCnt.toLocaleString() ?? 0,
        timeAvg: (sum_timeCnt / sum_theaterCnt).toFixed(1),
        seatCntAvg: (sum_totalSeatCnt / sum_timeCnt).toFixed(1),
        seatRateAvg: ((totalSoldSeatCnt / sum_totalSeatCnt) * 100).toFixed(1),
      });

      dataLength += 1;
      smallDataLength += 1;

      const mergeCell = `A${addIdx - smallDataLength + 2}:A${addIdx + 1}`;
      workSheet.mergeCells(mergeCell);

      this.addEmptyValue(workSheet, row);
    }

    workSheet.findRow(1).eachCell((cell, num) => {
      setBorder(workSheet, cell, 'thin');
      setFill(cell, 'e2efda');
      setFont2(cell, { bold: true });
      setCenter(cell);
    });
    workSheet.findRows(2, workSheet.lastRow.number).forEach((row) => {
      const checkCell = row.getCell('refineAreaNm');
      if (String(checkCell?.value ?? '') === '계') {
        row.eachCell((cell, num) => {
          if (num !== 1) {
            setFill(cell, 'ffe699');
          }
          setFont2(cell, { bold: true });
          setCenter(cell);
          setBorder(workSheet, cell, 'thin');
        });
      } else {
        header.forEach(({ key }) => {
          const cell = row.getCell(key);
          setBorder(workSheet, cell, 'thin');
          setCenter(cell);

          if (key === 'refineAreaNm') {
            setFont2(cell, { bold: true });
          }

          if (key === 'screenDt') {
            setFill(cell, 'e2efda');
            setFont2(cell, { bold: true });
            setCenter(cell);
          }
        });
      }
    });
  }

  async createDataByTheaterCompanySheet(workbook: ExcelJS.Workbook, dataByTheaterCompanyObj: any) {
    const workSheet: ExcelJS.Worksheet = setWorkSheet(workbook, '계열사별');
    const header = [
      { header: '상영일자', key: 'screenDt', columnKey: 'screenDt', width: 19 },
      { header: '계열사', key: 'theaterCompany', columnKey: 'theaterCompany' },
      { header: '극장수', key: 'theaterCnt', columnKey: 'theaterCnt' },
      { header: '상영회차', key: 'timeCnt', columnKey: 'timeCnt' },
      { header: '상영관수', key: 'hallCnt', columnKey: 'hallCnt' },
      { header: '총 좌석수', key: 'totalSeatCnt', columnKey: 'totalSeatCnt' },
      { header: '평균회차', key: 'timeAvg', columnKey: 'timeAvg' },
      { header: '평균좌석수', key: 'seatCntAvg', columnKey: 'seatCntAvg', width: 15 },
      { header: '평균 좌판율', key: 'seatRateAvg', columnKey: 'seatRateAvg', width: 15 },
    ];

    workSheet.columns = header;
    let dataLength = 0;
    let addIdx = 0;
    for await (const screenDt of Object.keys(dataByTheaterCompanyObj)) {
      let smallDataLength = 0;
      const screenData = dataByTheaterCompanyObj[screenDt];
      let sum_theaterCnt = 0;
      let sum_timeCnt = 0;
      let sum_hallCnt = 0;
      let sum_totalSeatCnt = 0;
      let sum_timeAvg = 0;
      let sum_remainSeatCnt = 0;

      for await (const theaterCompany of this.sortObjKeyAsc(screenData)) {
        const value = { ...screenData[theaterCompany], screenDt, theaterCompany };
        sum_theaterCnt += value.theaterCnt;
        sum_timeCnt += value.timeCnt;
        sum_hallCnt += value.hallCnt;
        sum_totalSeatCnt += value.totalSeatCnt;
        sum_timeAvg += value.timeAvg;
        sum_remainSeatCnt += value.totalRemainSeat;

        const soldSeatCnt = value.totalSeatCnt - value.totalRemainSeat;
        value.seatRateAvg = ((soldSeatCnt / value.totalSeatCnt) * 100).toFixed(1);
        workSheet.addRow({
          ...value,
          theaterCompany: `${value.theaterCompany.toUpperCase()} 계`,
          theaterCnt: value.theaterCnt.toLocaleString(),
          timeCnt: value.timeCnt.toLocaleString(),
          hallCnt: value.hallCnt.toLocaleString(),
          totalSeatCnt: value.totalSeatCnt.toLocaleString(),
          timeAvg: value.timeAvg ? value.timeAvg.toFixed(1) : '',
          seatCntAvg: value.seatCntAvg ? value.seatCntAvg.toFixed(1) : '',
        });
        addIdx++;
        dataLength += 1;
        smallDataLength += 1;
      }
      addIdx += 1;

      const totalSoldSeatCnt = sum_totalSeatCnt - sum_remainSeatCnt;

      const row = workSheet.addRow({
        screenDt,
        theaterCompany: '계',
        theaterCnt: sum_theaterCnt.toLocaleString() ?? 0,
        timeCnt: sum_timeCnt.toLocaleString() ?? 0,
        hallCnt: sum_hallCnt.toLocaleString() ?? 0,
        totalSeatCnt: sum_totalSeatCnt?.toLocaleString() ?? 0,
        timeAvg: (sum_timeCnt / sum_theaterCnt).toFixed(1),
        seatCntAvg: (sum_totalSeatCnt / sum_timeCnt).toFixed(1),
        seatRateAvg: ((totalSoldSeatCnt / sum_totalSeatCnt) * 100).toFixed(1),
      });
      dataLength += 1;
      smallDataLength += 1;

      const mergeCell = `A${addIdx - smallDataLength + 2}:A${addIdx + 1}`;
      workSheet.mergeCells(mergeCell);

      this.addEmptyValue(workSheet, row);
    }

    workSheet.findRow(1).eachCell((cell, num) => {
      setFill(cell, 'd9e1f2');
      setBorder(workSheet, cell, 'thin');
      setCenter(cell);
      setFont2(cell, { bold: true });
    });
    workSheet.findRows(2, dataLength + 2).forEach((row) => {
      const checkCell = row.getCell('theaterCompany');
      if (String(checkCell?.value ?? '') === '계') {
        row.eachCell((cell, num) => {
          if (num !== 1) {
            setFill(cell, 'ffe699');
          }

          setFont2(cell, { bold: true });
          setCenter(cell);
          setBorder(workSheet, cell, 'thin');
        });
      } else {
        header.forEach(({ key }) => {
          const cell = row.getCell(key);
          setBorder(workSheet, cell, 'thin');
          setCenter(cell);

          if (key === 'screenDt') {
            setFill(cell, 'd9e1f2');
            setFont2(cell, { bold: true });
          }
          if (key === 'theaterCompany') {
            setFont2(cell, { bold: true });
          }
        });
      }
    });
  }

  // 포맷 변환
  public refineType(findType: string) {
    findType = findType.replaceAll('(자막)', '').replaceAll('(더빙)', '');

    if (findType.includes('IMAX')) {
      return 'IMAX';
    } else if (findType.includes('4D')) {
      return '4D';
    } else if (findType.includes('3D')) {
      return '3D';
    } else if (findType.includes('DOLBY')) {
      return 'DOLBY';
    } else if (findType.includes('2D')) {
      return '2D';
    } else return findType;
  }

  public async createDataByFormat(workbook: ExcelJS.Workbook, compareDataForTheater: TScreenTable, crawlTime: string) {
    const workSheet: ExcelJS.Worksheet = setWorkSheet(workbook, '포맷별 요약표');

    const header = [
      { header: '상영일자', key: 'screenDt', columnKey: 'screenDt', width: 18 },
      { header: '포맷별', key: 'note', columnKey: 'note', width: 16 },
      { header: '계열사', key: 'theaterCompany', columnKey: 'theaterCompany' },
      { header: '극장수', key: 'theaterCnt', columnKey: 'theaterCnt', width: 10 },
      { header: '평균회차', key: 'timeAvg', columnKey: 'timeAvg', width: 9 },
      { header: '상영관수', key: 'hallCnt', columnKey: 'hallCnt', width: 13 },
      { header: '좌석수', key: 'totalSeatCnt', columnKey: 'totalSeatCnt', width: 15 },
      { header: '평균좌석수', key: 'seatCntAvg', columnKey: 'seatCntAvg', width: 11 },
    ];

    const headerRow = header.reduce((acc, { header, key }) => {
      acc[key] = header;
      return acc;
    }, {});

    workSheet.columns = header;

    const crawlRow = workSheet.addRow({
      screenDt: crawlTime,
    });
    workSheet.mergeCells('A2:H2');
    const crawlCell = workSheet.getCell('A2');

    setFont2(crawlCell, { bold: true, color: 'ff3a00' });

    const hallArray = [];
    const theaterArray = [];
    let refineData = ObjectUtil.newJson<TScreenTable>(compareDataForTheater).reduce(
      (acc, { screenDt, data: { data } }) => {
        if (!acc[screenDt]) {
          acc[screenDt] = {};
        }
        this.sortObjKeyAsc(data).forEach((theaterCompany) => {
          data[theaterCompany].forEach((d) => {
            let { note, type, theaterCompany, hall, theaterNm, totalSeatCnt, totalRemainSeat, totalTimeCnt } = d;
            let findType = type;
            let noteArray = [];
            noteArray.push(this.refineType(findType));

            if (note.includes('자막')) {
              // refineNote = '자막';
              noteArray.push('자막');
            } else if (note.includes('더빙')) {
              // refineNote = '더빙';
              noteArray.push('더빙');
            }
            noteArray.forEach((refineNote) => {
              if (!acc[screenDt][refineNote]) {
                acc[screenDt][refineNote] = {};
              }
              if (!acc[screenDt][refineNote][theaterCompany]) {
                acc[screenDt][refineNote][theaterCompany] = {
                  totalSeatCnt: 0,
                  totalRemainSeat: 0,
                  totalTimeCnt: 0,
                };
              }

              acc[screenDt][refineNote][theaterCompany] = {
                totalSeatCnt: totalSeatCnt + acc[screenDt][refineNote][theaterCompany].totalSeatCnt,
                totalRemainSeat: totalRemainSeat + acc[screenDt][refineNote][theaterCompany].totalRemainSeat,
                totalTimeCnt: totalTimeCnt + acc[screenDt][refineNote][theaterCompany].totalTimeCnt,
              };

              const refineHallNm = `${screenDt}_${refineNote}_${theaterCompany}_${theaterNm}_${hall}`;
              const refineTheaterNm = `${screenDt}_${refineNote}_${theaterCompany}_${theaterNm}`;

              if (!hallArray.includes(refineHallNm)) {
                hallArray.push(refineHallNm);
              }

              if (!theaterArray.includes(refineTheaterNm)) {
                theaterArray.push(refineTheaterNm);
              }
            });
          });
        });

        return acc;
      },
      {}
    );
    let firstRowNumber = 3;

    const addedHeaderRow = workSheet.addRow({
      ...headerRow,
    });
    let endDateRowArray = [];
    this.sortObjKeyAsc(refineData).forEach((screenDt, screenIdx) => {
      const dataByScreenDt = refineData[screenDt];

      let firstIdxByNoteIdx = 0;
      Object.keys(dataByScreenDt).forEach((note, noteIdx) => {
        const dataByNote = dataByScreenDt[note];

        let sumInfo = {
          theaterCnt: 0,
          timeCnt: 0,
          hallCnt: 0,
          totalSeatCnt: 0,
          totalRemainSeat: 0,
          timeAvg: 0,
          seatCntAvg: 0,
        };
        let firstIdx = 2;
        Object.keys(dataByNote).forEach((theaterCompany, dataByNoteIdx) => {
          const dataBytheaterCompany = dataByNote[theaterCompany];

          const { totalSeatCnt, totalRemainSeat, totalTimeCnt } = dataBytheaterCompany;
          const theaterCnt = theaterArray.filter((theaterNm) =>
            theaterNm.includes(`${screenDt}_${note}_${theaterCompany}`)
          ).length;
          const hallCnt = hallArray.filter((hall) => hall.includes(`${screenDt}_${note}_${theaterCompany}`)).length;
          let soldSeatCnt = totalSeatCnt - totalRemainSeat;

          const timeAvg = totalTimeCnt / theaterCnt;
          const seatCntAvg = (soldSeatCnt / totalSeatCnt) * 100;
          sumInfo = {
            ...sumInfo,
            theaterCnt: sumInfo.theaterCnt + theaterCnt,
            timeCnt: sumInfo.timeCnt + totalTimeCnt,
            hallCnt: sumInfo.hallCnt + hallCnt,
            totalRemainSeat: sumInfo.totalRemainSeat + totalRemainSeat,
            totalSeatCnt: sumInfo.totalSeatCnt + totalSeatCnt,
          };

          const addedRow = workSheet.addRow({
            screenDt,
            note,
            theaterCompany: theaterCompany.toUpperCase(),
            theaterCnt: theaterCnt.toLocaleString(),
            timeCnt: totalTimeCnt.toLocaleString(),
            hallCnt: hallCnt.toLocaleString(),
            totalSeatCnt: totalSeatCnt.toLocaleString(),
            timeAvg: timeAvg.toFixed(1),
            seatCntAvg: seatCntAvg.toFixed(1),
          });
          if (dataByNoteIdx === 0 && noteIdx === 0) {
            firstIdxByNoteIdx = addedRow.number;
          }
          if (dataByNoteIdx === 0) {
            firstIdx = addedRow.number;
          }
        });
        let sumSoldSeatCnt = sumInfo.totalSeatCnt - sumInfo.totalRemainSeat;

        const sumTimeAvg = sumInfo.timeCnt / sumInfo.theaterCnt;
        const sumSeatCntAvg = (sumSoldSeatCnt / sumInfo.totalSeatCnt) * 100;
        workSheet.addRow({
          note,
          screenDt,
          theaterCnt: sumInfo.theaterCnt.toLocaleString(),
          timeCnt: sumInfo.timeCnt.toLocaleString(),
          hallCnt: sumInfo.hallCnt.toLocaleString(),
          totalSeatCnt: sumInfo.totalSeatCnt.toLocaleString(),
          theaterCompany: '계',
          timeAvg: sumTimeAvg.toFixed(1),
          seatCntAvg: sumSeatCntAvg.toFixed(1),
        });

        const lastRowNumber = workSheet.lastRow.number;
        workSheet.mergeCells(`B${firstIdx}: B${lastRowNumber}`);
      });
      workSheet.mergeCells(`A${firstIdxByNoteIdx}:A${workSheet.lastRow.number}`);

      firstRowNumber = workSheet.lastRow.number;

      endDateRowArray.push(workSheet.lastRow.number);
    });

    workSheet.findRows(3, workSheet.lastRow.number).forEach((row) => {
      const noteCell = row.getCell('note');
      const checkCell = row.getCell('theaterCompany');

      if (!checkCell.value) return;

      // 공통
      header.forEach(({ key }) => {
        const cell = row.getCell(key);
        setBorder2(cell, { all: 'thin' });
        setFont2(cell, { size: 10, color: '#000' });

        if (key === 'theaterCompany' || key === 'screenDt') {
          setFont2(cell, { bold: true });
          setCenter(cell);
          setFill(cell, 'c5d9f1');
        }

        if (key === 'seatCntAvg') {
          setBorder2(cell, { right: 'medium' });
        }

        setCenter(cell);
      });
      if (checkCell.value === '계열사') {
        // 헤더일 때
        header.forEach(({ key }) => {
          const cell = row.getCell(key);
          setFont2(cell, { bold: true });
          setFill(cell, 'c5d9f1');
        });
      } else if (checkCell.value === '계') {
        header.forEach(({ key }) => {
          const cell = row.getCell(key);
          if (key !== 'theaterCompany' && key !== 'screenDt' && key !== 'note') {
            setRight(cell);
          }
          if (key !== 'screenDt') {
            fillCell(noteCell.value, cell);
          }
          setBold(cell, true);
        });
      } else {
        // 값일 때
        header.forEach(({ key }) => {
          const cell = row.getCell(key);
          if (key === 'theaterCompany' || key === 'note') {
            fillCell(noteCell.value, cell);
          } else if (key !== 'note') {
            setRight(cell);
          }
          if (key === 'screenDt') {
            setCenter(cell);
          }
        });
      }

      if (endDateRowArray.includes(row.number)) {
        header.forEach(({ key }) => {
          const cell = row.getCell(key);
          setBorder2(cell, { bottom: 'medium' });
        });
      }
    });

    /**
     * @description: 포맷별 요약표 시트에서 사용되는 함수로, 자막 / 더빙 등등의 값에 따른 색깔을 지정해주는 함수
     */
    function fillCell(value, cell) {
      switch (value) {
        case '자막':
          setFill(cell, 'eaf2df');
          break;
        case '더빙':
          setFill(cell, 'd9eff3');
          break;
        case 'IMAX':
          setFill(cell, 'e3e0ea');
          break;
        case '4D':
          setFill(cell, 'ffff99');
          break;
        case '3D':
          setFill(cell, 'd9e1f2');
          break;
        case 'DOLBY':
          setFill(cell, 'f7f9e1');
          break;
        case '2D':
          setFill(cell, 'ebf6a6');
          break;
      }
    }

    // 기존에는 header로 설정된 row는 1번째 줄로 고정된다.
    // 크롤링 날짜를 1번째 줄에 놓고, header의 설정된 row을 아래로 바꾸게 해주기 위해 아래와 같이 코드 작성하였음.
    workSheet.spliceRows(1, 1, []);
  }
}
