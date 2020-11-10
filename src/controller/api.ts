import {
  Inject,
  Controller,
  Post,
  Provide,
  Query,
  Get,
} from '@midwayjs/decorator';
import { Context } from 'egg';
import { UserService } from '../service/user';

const ExcelJS = require('exceljs');
const path = require('path');
const { MongoClient } = require('mongodb');

interface Employee {
  _id: string;
  name: string;
  englishName: string;
  position: string;
  subcontractor: string;
}

@Provide()
@Controller('/api')
export class APIController {
  @Inject()
  ctx: Context;

  @Inject()
  userService: UserService;

  @Post('/get_user')
  async getUser(@Query() uid) {
    const user = await this.userService.getUser({ uid });
    return { success: true, message: 'OK', data: user };
  }

  @Get('/employeesList')
  async employeesList() {
    const uri =
      'mongodb+srv://ZQH:1997121@emploee.pjlnq.azure.mongodb.net/EmployeeManage?retryWrites=true&w=majority';
    const client = new MongoClient(uri, { useUnifiedTopology: true });
    const employees: any = [];
    async function run() {
      try {
        await client.connect();
        const database = client.db('EmployeeManage');
        const collection = database.collection('employees');
        const cursor = collection.find();
        await cursor.forEach((item: any) => {
          employees.push(item);
        });
      } finally {
        // Ensures that the client will close when you finish/error
        await client.close();
      }
    }
    await run().catch(console.dir);

    return {
      employees: employees,
    };
  }

  @Post('/createXLSX')
  async createXLSX() {
    const body = this.ctx.request.body;
    const employees: Employee[] = body.selected;
    const selectedDate: string[] = body.selectedDate;
    const holiday: string = body.holiday;
    const days: number = body.days;
    const data: any[] = [];
    async function getData(
      employee: Employee,
      date: string[],
      holiday: string,
      days: number
    ) {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(
        path.join(__dirname, '../../src/XLSX/report.xlsx')
      );
      const worksheet = workbook.getWorksheet(1);
      worksheet.getCell('A4').value = employee.name;
      worksheet.getCell('C4').value = employee._id;
      worksheet.getCell('C5').value = holiday;
      worksheet.getCell('E4').value = employee.position;
      worksheet.getCell('C7').value = date[0];
      worksheet.getCell('E7').value = date[1];
      worksheet.getCell('F7').value = days;
      worksheet.getCell('C14').value = employee.name;
      data.push({
        name: employee.name,
        value: JSON.stringify(await workbook.xlsx.writeBuffer()),
      });
    }
    if (employees[0]) {
      for (const employee of employees) {
        await getData(employee, selectedDate, holiday, days);
      }
    }
    return { data };
  }
}
