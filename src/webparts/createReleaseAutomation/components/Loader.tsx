import * as React from "react";
import "./Loader.css";
const Loader = () => {
  return (
    <div className="loadercontainer">
      <div className="loading">
        <span></span>
        <span></span>
        <span></span>
        <span></span>
        <span></span>
      </div>
      <p style={{ margin: 0 }}>
      It may take 3-5 minutes; please refrain from refreshing the page.
      </p>
    </div>
  );
};
export default Loader;
