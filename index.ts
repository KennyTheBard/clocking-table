import xlsx from 'node-xlsx';
import moment from 'moment';
import * as fs from 'fs';

const weeks: any[][] = [['Week', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday', 'Weekly total']];
const today = moment();
const day = today.clone().startOf('month');
const lastDay = today.clone().endOf('month');
let rowNumber = 2;
while (day.isBefore(lastDay)) {
    const firstDayOfWeek = day.clone().startOf('isoWeek');
    const lastDayOfWeek = day.clone().endOf('isoWeek');
    const week: any[] = [
        `${firstDayOfWeek.format('DD.MM')} - ${lastDayOfWeek.format('DD.MM')}`
    ];
    for (let day = firstDayOfWeek; day.isSameOrBefore(lastDayOfWeek); day.add(1, 'day')) {
        week.push(day.isSame(today, 'month') ? 0 : 'X')
    }
    week.push({f: `=SUM(B${rowNumber}:H${rowNumber})`})
    weeks.push(week);
    day.add(1, 'week');
    rowNumber++;
}
weeks.push(['', '', '', '','','','','',{f: `=SUM(I2:I${rowNumber - 1})`}])

const sheetOptions = {'!cols': [{wch: 15}, {wch: 10}, {wch: 10}, {wch: 10}, {wch: 10}, {wch: 10}, {wch: 10}, {wch: 10}, {wch: 10}]};
const buffer = xlsx.build([{name: 'Sheet', data: weeks, options: sheetOptions}]);
fs.writeFileSync('./table.xlsx', buffer)

// TODO: add colon widths

// TODO: add color maybe?