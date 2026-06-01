import React, { useState } from 'react'
import * as XLSX from 'xlsx'
import {
  BarChart,
  Bar,
  LineChart,
  Line,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  CartesianGrid,
} from 'recharts'

export default function Dashboard() {
  const [salesData, setSalesData] = useState([])
  const [branchData, setBranchData] = useState([])
  const [paymentData, setPaymentData] = useState([])
  const [kpis, setKpis] = useState({
    totalSales: 0,
    totalCash: 0,
    totalVisa: 0,
    invoices: 0,
    topBranch: '-',
    lowBranch: '-',
  })

  const handleFileUpload = (e) => {
    const file = e.target.files[0]
    if (!file) return

    const reader = new FileReader()

    reader.onload = (evt) => {
      const data = evt.target.result
      const workbook = XLSX.read(data, { type: 'binary' })
      const sheetName = workbook.SheetNames[0]
      const worksheet = workbook.Sheets[sheetName]
      const jsonData = XLSX.utils.sheet_to_json(worksheet)

      console.log("[v0] Raw data from Excel:", jsonData)
      console.log("[v0] First row columns:", jsonData[0] ? Object.keys(jsonData[0]) : "No data")
      
      // Log exact column names with their character codes for debugging
      if (jsonData[0]) {
        Object.keys(jsonData[0]).forEach(key => {
          console.log(`[v0] Column: "${key}" | Length: ${key.length} | Values sample:`, jsonData[0][key])
        })
      }

      // Helper function to parse numbers that may be in accounting format (1,234.00) or negative (1,234.00)
      const parseNumber = (value) => {
        if (value === undefined || value === null || value === '' || value === '-') {
          return 0
        }
        // If already a number, return it
        if (typeof value === 'number') {
          return value
        }
        // Convert to string and handle accounting format
        let str = String(value).trim()
        // Check if negative (wrapped in parentheses)
        const isNegative = str.startsWith('(') && str.endsWith(')')
        if (isNegative) {
          str = str.slice(1, -1) // Remove parentheses
        }
        // Remove commas and spaces
        str = str.replace(/,/g, '').replace(/\s/g, '')
        const num = parseFloat(str)
        if (isNaN(num)) return 0
        return isNegative ? -num : num
      }

      // Helper function to find column value with fuzzy matching
      const getColumnValue = (row, possibleNames, defaultValue = 0) => {
        // First try exact match
        for (const name of possibleNames) {
          if (row[name] !== undefined && row[name] !== null && row[name] !== '') {
            return row[name]
          }
        }
        // Then try partial match (column contains the search term)
        const rowKeys = Object.keys(row)
        for (const name of possibleNames) {
          for (const key of rowKeys) {
            // Normalize both strings for comparison
            const normalizedKey = key.trim().replace(/\s+/g, ' ')
            const normalizedName = name.trim().replace(/\s+/g, ' ')
            if (normalizedKey.includes(normalizedName) || normalizedName.includes(normalizedKey)) {
              if (row[key] !== undefined && row[key] !== null && row[key] !== '') {
                return row[key]
              }
            }
          }
        }
        return defaultValue
      }

      let totalSales = 0
      let totalCash = 0
      let totalVisa = 0

      const salesByDay = {}
      const salesByBranch = {}

      // Possible column name variations - including partial matches
      const dateColumns = ['التاريخ', 'تاريخ', 'Date', 'date', 'اليوم', 'يوم']
      const branchColumns = ['الفرع', 'فرع', 'Branch', 'branch', 'المتجر', 'متجر', 'Store']
      const invoiceColumns = ['اجمالي الفاتورة', 'إجمالي الفاتورة', 'اجمالى الفاتورة', 'إجمالى الفاتورة', 'الإجمالي', 'الاجمالي', 'إجمالي', 'اجمالي', 'اجمالى', 'المبيعات', 'مبيعات', 'Total', 'total', 'Sales', 'sales', 'Amount', 'amount', 'المبلغ', 'مبلغ', 'القيمة', 'قيمة', 'الفاتورة']
      const cashColumns = ['الكاش', 'كاش', 'Cash', 'cash', 'نقدي', 'النقدي', 'نقد']
      const visaColumns = ['الفيزا', 'فيزا', 'Visa', 'visa', 'بطاقة', 'كارت', 'Card', 'card', 'شبكة']

      jsonData.forEach((row, index) => {
        const date = getColumnValue(row, dateColumns, '-') || '-'
        const branch = getColumnValue(row, branchColumns, 'غير محدد') || 'غير محدد'

        const invoice = parseNumber(getColumnValue(row, invoiceColumns, 0))
        const cash = parseNumber(getColumnValue(row, cashColumns, 0))
        const visa = parseNumber(getColumnValue(row, visaColumns, 0))

        if (index === 0) {
          console.log("[v0] First row parsed values:", { date, branch, invoice, cash, visa })
        }

        totalSales += invoice
        totalCash += cash
        totalVisa += visa

        salesByDay[date] = (salesByDay[date] || 0) + invoice
        salesByBranch[branch] = (salesByBranch[branch] || 0) + invoice
      })

      console.log("[v0] Totals calculated:", { totalSales, totalCash, totalVisa })

      const daily = Object.keys(salesByDay).map((key) => ({
        day: key,
        sales: salesByDay[key],
      }))

      const branches = Object.keys(salesByBranch).map((key) => ({
        branch: key,
        sales: salesByBranch[key],
      }))

      const sortedBranches = [...branches].sort((a, b) => b.sales - a.sales)

      setKpis({
        totalSales,
        totalCash,
        totalVisa,
        invoices: jsonData.length,
        topBranch: sortedBranches[0]?.branch || '-',
        lowBranch: sortedBranches[sortedBranches.length - 1]?.branch || '-',
      })

      setSalesData(daily)
      setBranchData(branches)

      setPaymentData([
        {
          name: 'كاش',
          value: totalCash,
        },
        {
          name: 'فيزا',
          value: totalVisa,
        },
      ])
    }

    reader.readAsBinaryString(file)
  }

  return (
    <div className="min-h-screen bg-[#0f172a] text-white p-6" dir="rtl">
      <div className="flex flex-col lg:flex-row justify-between items-center mb-8 gap-4">
        <div>
          <h1 className="text-4xl font-bold mb-2">لوحة تحكم المبيعات</h1>
          <p className="text-slate-400">Dashboard احترافية مرتبطة بملف Excel</p>
        </div>

        <label className="bg-blue-600 hover:bg-blue-700 px-5 py-3 rounded-2xl cursor-pointer font-bold transition">
          تحميل ملف Excel
          <input
            type="file"
            accept=".xlsx,.xls"
            className="hidden"
            onChange={handleFileUpload}
          />
        </label>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-6 gap-4 mb-8">
        <Card title="إجمالي المبيعات" value={formatNumber(kpis.totalSales)} />
        <Card title="إجمالي الكاش" value={formatNumber(kpis.totalCash)} />
        <Card title="إجمالي الفيزا" value={formatNumber(kpis.totalVisa)} />
        <Card title="عدد الفواتير" value={kpis.invoices} />
        <Card title="أعلى فرع" value={kpis.topBranch} />
        <Card title="أقل فرع" value={kpis.lowBranch} />
      </div>

      <div className="grid grid-cols-1 xl:grid-cols-2 gap-6 mb-8">
        <div className="bg-[#1e293b] rounded-3xl p-6 border border-slate-700 shadow-lg">
          <h2 className="text-2xl font-bold mb-6">المبيعات اليومية</h2>

          <ResponsiveContainer width="100%" height={350}>
            <LineChart data={salesData}>
              <CartesianGrid strokeDasharray="3 3" stroke="#334155" />
              <XAxis dataKey="day" stroke="#94a3b8" />
              <YAxis stroke="#94a3b8" />
              <Tooltip />
              <Line
                type="monotone"
                dataKey="sales"
                stroke="#3b82f6"
                strokeWidth={3}
              />
            </LineChart>
          </ResponsiveContainer>
        </div>

        <div className="bg-[#1e293b] rounded-3xl p-6 border border-slate-700 shadow-lg">
          <h2 className="text-2xl font-bold mb-6">مبيعات الفروع</h2>

          <ResponsiveContainer width="100%" height={350}>
            <BarChart data={branchData}>
              <CartesianGrid strokeDasharray="3 3" stroke="#334155" />
              <XAxis dataKey="branch" stroke="#94a3b8" />
              <YAxis stroke="#94a3b8" />
              <Tooltip />
              <Bar dataKey="sales" fill="#3b82f6" radius={[10, 10, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="grid grid-cols-1 xl:grid-cols-3 gap-6">
        <div className="bg-[#1e293b] rounded-3xl p-6 border border-slate-700 shadow-lg xl:col-span-1">
          <h2 className="text-2xl font-bold mb-6">الكاش والفيزا</h2>

          <ResponsiveContainer width="100%" height={300}>
            <PieChart>
              <Pie
                data={paymentData}
                dataKey="value"
                outerRadius={100}
                label
              >
                <Cell fill="#3b82f6" />
                <Cell fill="#22c55e" />
              </Pie>
              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
        </div>

        <div className="bg-[#1e293b] rounded-3xl p-6 border border-slate-700 shadow-lg xl:col-span-2 overflow-auto">
          <h2 className="text-2xl font-bold mb-6">تفاصيل المبيعات</h2>

          <table className="w-full text-right">
            <thead>
              <tr className="border-b border-slate-700 text-slate-300">
                <th className="p-3">اليوم</th>
                <th className="p-3">المبيعات</th>
              </tr>
            </thead>

            <tbody>
              {salesData.map((item, index) => (
                <tr
                  key={index}
                  className="border-b border-slate-800 hover:bg-slate-800 transition"
                >
                  <td className="p-3">{item.day}</td>
                  <td className="p-3">{formatNumber(item.sales)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}

function Card({ title, value }) {
  return (
    <div className="bg-[#1e293b] rounded-3xl p-5 border border-slate-700 shadow-lg">
      <h2 className="text-slate-400 mb-3 text-sm">{title}</h2>
      <p className="text-2xl font-bold">{value}</p>
    </div>
  )
}

function formatNumber(num) {
  return new Intl.NumberFormat('ar-EG').format(num || 0)
}
