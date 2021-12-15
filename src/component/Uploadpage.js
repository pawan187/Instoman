import React, { useState } from 'react';
import axios from 'axios';
export default () => {
  axios
    .get(
      'https://e4f94.sse.codesandbox.io/mtc'
      //  {
      //   firstName: 'Finn',
      //   lastName: 'Williams'
      // }
    )
    .then(
      (response) => {
        console.log('server is active');
      },
      (error) => {
        console.log(error);
      }
    );
  const [docuementList, setdocuementList] = useState([]);

  const [files, setFiles] = useState([]);

  let [color, setColor] = useState('#ffffff');
  const [id,setid] = useState()
  const uploadFiles = (e) => {
    e.preventDefault();
    var data = new FormData()
    let newList = [];
    for (let i = 0; i < files.length; i++) {
      newList = newList.concat({
        name: files[i].name,
        upload: false,
      });

      data.append('images',files[i],files[i].name)

    }
    setdocuementList(docuementList.concat(newList));

    fetch('https://e4f94.sse.codesandbox.io/mtc/uploadBulkImage', {
      method: 'post',
      body: data
    }).then((res) => {
      console.log(res.data.files);
      
    });
    console.log(docuementList.concat(newList));
  };
  return (
    <div class="container">
      <h1>Form to upload documents</h1>
      <form onSubmit={uploadFiles}>
        <div class="input-group">
          <input
            type="file"
            class="form-control"
            id="images"
            name="images"
            aria-describedby="inputGroupFileAddon04"
            aria-label="Upload"
            onChange={(e) => {
              setFiles(e.target.files);
              console.log(e.target.files.length);
            }}
            multiple
          ></input>
          <button
            class="btn btn-outline-secondary"
            type="submit"
            id="inputGroupFileAddon04"
          >
            Button
          </button>
        </div>
      </form>
      <div class="container">
        <p>
          Table to get the status of document like uploaded, arm no, extraction,
          comparison, final report generation
        </p>
        <ul class="list-group">
          <li class="list-group-item">
            <div class="container">
              <div class="card bg-ligth">
                <div class="row">
                  <div class="col">Document name</div>
                  <div class="col">Uploaded</div>
                </div>
              </div>
              {docuementList.map((element, index) => {
                return (
                  <div class="card">
                    <div class="row" key={index}>
                      <div class="col">{element.name}</div>
                      <div class="col">
                        {element.upload
                          ? 'successful'
                          : // <ClipLoader
                            //   color={color}
                            //   loading={!element.upload}
                            //   css={override}
                            //   size={150}
                            // />
                            'successful'}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </li>
        </ul>
      </div>
      <div class="container">
        {docuementList.length > 0 ? (
          
          <a href={"/reports:" + id }> Run Bot </a>
        ) : (
          // <ClipLoader
          //   color={color}
          //   loading={!element.upload}
          //   css={override}
          //   size={150}
          // />
          'please uplaod documents'
        )}
      </div>
    </div>
  );
};
