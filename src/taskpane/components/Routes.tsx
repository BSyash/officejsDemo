import React from "react";
import { Switch, BrowserRouter, Route } from "react-router-dom";
import App from "./App";
import DashBoard from "./DashBoard";

const Routes = () => {
  return (
      <BrowserRouter>
        <Switch>
          <Route path={"/"} component={App} />
          <Route path={"/dashboard"} component={DashBoard} />
        </Switch>
      </BrowserRouter>
  );
};

export default Routes;
