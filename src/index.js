import React, { useState } from "react";
import ReactDOM from "react-dom";
import ExcelImport from "./ExcelImport";
import ExcelExampleExport from "./ExcelExampleExport";

import "./styles.css";

const App = () => {
  const [data, setData] = useState([]);

  const createRequests = () => {
    console.log(data);
  };

  const worksheets = [
    {
      name: "Requests",
      columns: [
        { label: "Full Name", value: "name" },
        { label: "Email", value: "email" },
        { label: "Template", value: "template" }
      ],
      data: [
        {
          name: "Bob Ross",
          email: "boss_ross@gmail.com",
          template: "Accounts Receivables"
        }
      ]
    }
  ];
  return (
    <div>
      <ExcelExampleExport filename="requests.xlsx" worksheets={worksheets} />
      <ExcelImport uploadHandler={setData} />
      <button onClick={createRequests}>Create</button>
    </div>
  );
};

const rootElement = document.getElementById("root");
ReactDOM.render(<App />, rootElement);
