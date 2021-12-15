import React, { useState } from 'react';

import Home from '../component/Uploadpage.js';
import Reports from '../component/Reports';
import NavBar from '../component/NavBar';
import TCList from '../component/TCList'
// import Error from '../component/Error';
import ReactDOM from 'react-dom';
import {
  Router,
  BrowserRouter,
  Route,
  Switch,
  Link,
  NavLink,
} from 'react-router-dom';

import { createBrowserHistory } from 'history';

export const history = createBrowserHistory();

export default () => {
  return (
    <Router history={history}>
      <div>
        <NavBar />

        <Switch>
          <Route path='/' component={TCList} exact={true}/>
          <Route path="/Home" component={Home}  />
          <Route path="/reports:id" component={Reports} />
          {/* <Route component={Error} /> */}
        </Switch>
      </div>
    </Router>
  );
};
