import type { Lesson } from "./types";

export const finance: Lesson[] = [
  {
    slug: "pmt",
    title: "PMT - calculate a loan repayment",
    track: "finance",
    level: "Intermediate",
    summary:
      "PMT returns the fixed periodic payment for a loan. Get the rate and periods into the same unit and it is almost impossible to get wrong.",
    objectives: [
      "Convert an annual rate to a periodic rate",
      "Understand why the result is negative",
      "Build a payment sensitivity table",
    ],
    syntax: "=PMT(rate; nper; pv; [fv]; [type])",
    steps: [
      {
        title: "Match the units",
        detail:
          "Monthly payments need a monthly rate (annual/12) and a number of months (years*12). Mixing them is the number-one error.",
      },
      {
        title: "Write it",
        detail: "£250 000 over 25 years at 5.5%.",
        formula: "=PMT(5.5%/12;25*12;250000)",
      },
      {
        title: "Why negative?",
        detail:
          "Excel uses cash-flow signs: money leaving you is negative. Put a minus in front for a positive display.",
        formula: "=-PMT(B1/12;B2*12;B3)",
      },
      {
        title: "Payments at the start of the period",
        detail: "type = 1 for annuity-due (rent paid in advance), 0 or omitted for the end.",
        formula: "=PMT(rate;nper;pv;0;1)",
      },
      {
        title: "Leave a balloon balance",
        detail: "fv is what remains owing at the end.",
        formula: "=-PMT(5%/12;60;30000;-5000)",
      },
      {
        title: "Sensitivity table",
        detail:
          "Put rates down a column and terms across a row, then Data → What-If Analysis → Data Table to see every combination at once.",
      },
    ],
    example: {
      caption: "Inputs to lay out on the sheet",
      headers: ["Cell", "Input", "Example"],
      rows: [
        ["B1", "Annual rate", "5.5%"],
        ["B2", "Years", "25"],
        ["B3", "Loan amount", "250 000"],
        ["B4", "Monthly payment", "=-PMT(B1/12;B2*12;B3) → 1 535.19"],
      ],
    },
    practice: {
      task: "Car loan: 18 000 over 5 years at 7.9% annual. Monthly payment?",
      answer: "=-PMT(7.9%/12;5*12;18000) → about 364",
    },
    quiz: [
      {
        q: "For a monthly payment on an 6% annual loan you use rate =",
        options: ["6%", "6%/12", "6%*12", "0.06/365"],
        answer: 1,
        explain: "The rate must be per period, so annual ÷ 12.",
      },
    ],
  },
  {
    slug: "ppmt-ipmt",
    title: "PPMT & IPMT - split a payment into capital and interest",
    track: "finance",
    level: "Advanced",
    summary:
      "PPMT gives the principal part of a chosen payment and IPMT gives the interest part. Together they build a full amortisation schedule.",
    objectives: [
      "Split any instalment",
      "Build a 12-row (or 360-row) schedule",
      "Prove that PPMT + IPMT = PMT",
    ],
    syntax: "=PPMT(rate; per; nper; pv; [fv]; [type])   =IPMT(rate; per; nper; pv; [fv]; [type])",
    steps: [
      {
        title: "The extra argument is 'per'",
        detail: "per is which payment number you are looking at - 1 for the first month.",
        formula: "=IPMT(5.5%/12;1;300;250000)",
      },
      { title: "Principal of the same payment", detail: "", formula: "=PPMT(5.5%/12;1;300;250000)" },
      {
        title: "They always add up",
        detail: "PPMT(per) + IPMT(per) = PMT for every period. Use it as a check row in your model.",
      },
      {
        title: "Build the schedule",
        detail:
          "Put the period number in column A (1, 2, 3 …), then lock the rate/term/loan cells with $ and copy the formulas down.",
        formula:
          "B2: =-IPMT($B$1/12;A2;$B$2*12;$B$3)   C2: =-PPMT($B$1/12;A2;$B$2*12;$B$3)   D2: =D1-C2",
      },
      {
        title: "Spill the whole schedule in one formula (365)",
        detail: "SEQUENCE supplies every period number at once.",
        formula: "=-PPMT(5.5%/12;SEQUENCE(300);300;250000)",
      },
      {
        title: "Total interest paid",
        detail: "The number that makes people overpay their mortgage.",
        formula: "=-SUM(IPMT(5.5%/12;SEQUENCE(300);300;250000))",
      },
    ],
    practice: {
      task: "How much interest is in payment 12 of a 200 000, 20-year, 4.5% loan?",
      answer: "=-IPMT(4.5%/12;12;240;200000)",
    },
    quiz: [
      {
        q: "Early in a repayment loan, most of each payment is:",
        options: ["Principal", "Interest", "Fees", "Equal split"],
        answer: 1,
        explain: "Interest is charged on a big outstanding balance, so it dominates early payments.",
      },
    ],
  },
  {
    slug: "pv",
    title: "PV - what a future stream is worth today",
    track: "finance",
    level: "Intermediate",
    summary:
      "PV discounts future payments back to today's money - the basis of lease valuation, annuities and 'should I take the lump sum?' decisions.",
    objectives: ["Value an annuity", "Value a single future amount", "Use PV to check a loan"],
    syntax: "=PV(rate; nper; pmt; [fv]; [type])",
    steps: [
      { title: "Value a stream of payments", detail: "1 000 a month for 10 years at 6% annual.", formula: "=-PV(6%/12;120;1000)" },
      { title: "Value a single future sum", detail: "Set pmt to 0 and use fv.", formula: "=-PV(6%;10;0;50000)" },
      {
        title: "Check a loan",
        detail:
          "PV of the payments should equal the loan amount - a neat self-test for a model.",
      },
      {
        title: "Choosing the discount rate",
        detail:
          "Use your cost of capital or the return you could get elsewhere. The answer is only as good as this assumption.",
      },
      { title: "Uneven cash flows", detail: "PV assumes equal payments; use NPV or XNPV when they vary.", formula: "=NPV(8%;B2:B11)+B1" },
    ],
    practice: {
      task: "A lease costs 2 500 per quarter for 5 years. At 8% annual, what is it worth today?",
      answer: "=-PV(8%/4;5*4;2500)",
    },
    quiz: [
      {
        q: "Which function handles uneven cash flows?",
        options: ["PV", "PMT", "NPV", "FV"],
        answer: 2,
        explain: "NPV (or XNPV with dates) discounts a list of different amounts.",
      },
    ],
  },
  {
    slug: "fv",
    title: "FV - what your savings will be worth",
    track: "finance",
    level: "Intermediate",
    summary:
      "FV projects a savings plan or investment forward, combining a starting lump sum with regular contributions and compound growth.",
    objectives: ["Project regular savings", "Combine a lump sum and contributions", "Compare monthly vs annual compounding"],
    syntax: "=FV(rate; nper; pmt; [pv]; [type])",
    steps: [
      { title: "Regular saving", detail: "200 a month for 20 years at 6%.", formula: "=-FV(6%/12;20*12;200)" },
      { title: "Add a starting balance", detail: "pv is negative because you pay it in.", formula: "=-FV(6%/12;240;200;-5000)" },
      { title: "Contribute at the start of each month", detail: "type = 1 adds one extra month of growth to every payment.", formula: "=-FV(6%/12;240;200;-5000;1)" },
      {
        title: "Reverse the question",
        detail: "How much must you save each month to reach 100 000? PMT answers it.",
        formula: "=-PMT(6%/12;240;0;100000)",
      },
      { title: "Show the effect of inflation", detail: "Use a real rate: (1+nominal)/(1+inflation)-1." },
    ],
    practice: {
      task: "You save 300 per month for 15 years at 5%. Final value?",
      answer: "=-FV(5%/12;15*12;300) → about 80 200",
    },
    quiz: [
      {
        q: "Why is pmt entered as a negative number?",
        options: [
          "A bug",
          "Cash paid out is negative in Excel's sign convention",
          "To make the result smaller",
          "It is not",
        ],
        answer: 1,
        explain: "Money leaving your pocket is a negative cash flow.",
      },
    ],
  },
  {
    slug: "mortgage-calculator",
    title: "Project: build an Excel Mortgage Calculator",
    track: "finance",
    level: "Advanced",
    summary:
      "Put PMT, PPMT, IPMT and a data table together into a complete, sharable mortgage model with an overpayment scenario.",
    objectives: [
      "Lay out an input block",
      "Build the summary and the schedule",
      "Add a rate sensitivity table and an overpayment comparison",
    ],
    steps: [
      {
        title: "1. Input block (shade it blue)",
        detail:
          "B1 Loan amount 250000, B2 Annual rate 5.5%, B3 Term years 25, B4 Payments per year 12, B5 Extra monthly overpayment 0.",
      },
      {
        title: "2. Derived values",
        detail: "B6 rate per period =B2/B4, B7 number of periods =B3*B4.",
      },
      {
        title: "3. The core payment",
        detail: "",
        formula: "B8: =-PMT(B6;B7;B1)",
      },
      {
        title: "4. Summary figures",
        detail: "Total paid and total interest tell the real story.",
        formula: "B9: =B8*B7      B10: =B9-B1",
      },
      {
        title: "5. The schedule",
        detail:
          "A13 headers: Period, Opening, Payment, Interest, Principal, Overpayment, Closing. A14 = 1, then =A14+1 down to the term.",
        formula:
          "B14: =$B$1  C14: =$B$8  D14: =B14*$B$6  E14: =C14-D14  F14: =$B$5  G14: =B14-E14-F14  then B15: =G14",
      },
      {
        title: "6. Stop when the loan is repaid",
        detail: "Wrap each row so it goes blank once the balance hits zero.",
        formula: "=IF(B14<=0;\"\";…)",
      },
      {
        title: "7. Rate sensitivity",
        detail:
          "List rates 3%-8% down a column next to a formula =B8, select the block and use Data → What-If Analysis → Data Table with the column input cell = B2.",
      },
      {
        title: "8. Overpayment saving",
        detail:
          "Count the remaining periods and compare interest. =COUNTIF(G14:G500;\">0\") gives the real term; the interest saved is the difference in SUM(D:D).",
      },
      {
        title: "9. Finish it properly",
        detail:
          "Format currency, freeze panes at row 14, protect the formula cells (Ctrl + 1 → Protection) and unlock only the blue input cells, then Review → Protect Sheet.",
      },
    ],
    example: {
      caption: "Expected result for the default inputs",
      headers: ["Figure", "Value"],
      rows: [
        ["Monthly payment", "1 535.19"],
        ["Total paid over 25 years", "460 557"],
        ["Total interest", "210 557"],
        ["With 200/month overpayment", "Repaid in ~21 years, saves ~46 000 interest"],
      ],
    },
    practice: {
      task: "Add a field for an arrangement fee added to the loan and show its lifetime cost.",
      answer:
        "Add B11 Fee. Change the payment to =-PMT(B6;B7;B1+B11) and compare total interest with and without the fee.",
    },
    quiz: [
      {
        q: "Which tool builds the rate sensitivity grid?",
        options: ["Goal Seek", "Data Table", "Solver", "Scenario Manager"],
        answer: 1,
        explain: "Data → What-If Analysis → Data Table recalculates the model for each input value.",
      },
    ],
  },
];
