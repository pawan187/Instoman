import React, { useState } from 'react';

export default () => {
  const [docuementList, setdocuementList] = useState([
    {
      name: 'TC-0468-7986',
      arm: '202215421_R0',
      upload: false,
      extraction: false,
      comparison: false,
      final_report: '/xyz.xlsx',
    },
  ]);
  const [counter,setcounter] = useState(60)

  const [files, setFiles] = useState([]);

  let [color, setColor] = useState('#ffffff');

  const uploadFiles = () => {
    let newList = [];
    for (let i = 0; i < files.length; i++) {
      newList = newList.concat({
        name: files[i].name,
        upload: false,
        extraction: false,
        comparison: false,
        report: 'path of final report',
      });
    }
    setdocuementList(docuementList.concat(newList));
    console.log(docuementList.concat(newList));
  };

  function sleep (time) {
    return new Promise((resolve) => setTimeout(resolve, time));
  }
  
  

  return (
    <div class="container">
      <p>
        Table to get the status of document like uploaded, arm no, extraction,
        comparison, final report generation
      </p>
      <ul class="list-group">
        <li class="list-group-item">
          <div class="container">
            <div class="row border text-light bg-primary">
              <div class="col">TC document</div>
              <div class="col">Arm</div>
              <div class="col">Extraction</div>
              <div class="col">comparison</div>
              <div class="col">final report</div>
            </div>
            {docuementList.map((element, index) => {
              return (
                <div class="row border" key={index}>
                  <div class="col">{element.name}</div>
                  <div class="col">
                    {element.arm ? (
                      element.arm
                    ) : (
                      // <ClipLoader
                      //   color={color}
                      //   loading={!element.upload}
                      //   css={override}
                      //   size={150}
                      // />
                      <div class="spinner-border text-primary" role="status">
                        <span class="sr-only"></span>
                      </div>
                    )}
                  </div>
                  <div class="col">
                    {element.extraction ? (
                      'successful'
                    ) : (
                      <div class="spinner-border text-primary" role="status">
                        <span class="sr-only"></span>
                      </div>
                    )}
                  </div>
                  <div class="col">
                    {element.comparison ? (
                      'successful'
                    ) : (
                      <div class="spinner-border text-primary" role="status">
                        <span class="sr-only"> { counter + 's'} </span>
                      </div>
                    )}
                  </div>
                  <div class="col">
                    {element.final_report ? (
                      <a href={element.final_report}class='btn'>
                      {element.final_report}
                      </a>
                    ) : (
                      <div class="spinner-grow text-primary" role="status">
                        <span class="sr-only"></span>
                      </div>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </li>
      </ul>
    </div>
  );
};
