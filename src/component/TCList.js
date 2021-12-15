import React from 'react';

export default () => {
  const vendor = [
    'President',
    'CHW',
    'Ratanamani',
    'AMNS',
    'ISMT',
    'Chandan Steel',
  ];
  return (
    <div>
      <div class="container">
        <h1 class="nav justify-content-center">Please select a vendor </h1>
      </div>
      <div class="container border">
        <div class="row">
          {vendor.map((element, index) => {
            return (
              <div class="col" key={index}>
                <div class="card">
                  {/* <img src="..." class="card-img-top" alt="..."></img> */}
                  <div class="card-body">
                    <h5 class="card-title">{element}</h5>
                    {/* <p class="card-text">
                Some quick example text to build on the card title and make up
                the bulk of the card's content.
              </p> */}
                    <a href="/Home" class="btn btn-primary">
                      Do Inspection
                    </a>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
};
