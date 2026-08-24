// src/ProtectedRoute.js
import React from "react";
import { Route, Redirect } from "react-router-dom";
import { isAuthenticated, getCurrentUser } from "./auth";
import Navbar from "./Navbar";

const ProtectedRoute = ({ component: Component, adminOnly = false, ...rest }) => {
  const currentUser = getCurrentUser();
  const isAdmin = (currentUser?.role || "").trim().toLowerCase() === "admin";

  return (
    <Route
      {...rest}
      render={(props) => {
        if (!isAuthenticated()) {
          return (
            <Redirect
              to={{
                pathname: "/",
                state: { from: props.location }
              }}
            />
          );
        }

        if (adminOnly && !isAdmin) {
          return <Redirect to="/dashboard" />;
        }

        return (
          <>
            <Navbar />
            <Component {...props} />
          </>
        );
      }}
    />
  );
};

export default ProtectedRoute;
