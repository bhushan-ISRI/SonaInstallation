import * as React from 'react';
import ApproverDashboard from './ApproverDashboard';
import "./usersite.scss";
import UserDashboard from "./UserDashboard";
import APperformerDashboard from './APperformerDashboard';
import { IDashboardProps } from './IDashboardProps';

import ApLogo from '../assets/ApDashboard.png';
import UserLogo from '../assets/UserDashboard.png';
import ApproverLogo from '../assets/ApproverDashboard.png';

export default function CapexDasboard(props: IDashboardProps) {

  const [page, setPage] = React.useState<string>("home");

  const navigate = (pageName: string) => {
    setPage(pageName);
  };

  return (
    <div>

      {page === "home" && (
        <div className="main-container">
          <div className="headSheet"><h2>Installation Commission</h2></div>
          <section className='hero'>
            <div className="overlay"></div>
            <div className="hero-content">
              <div className='card-container'>

                <div className="infoCard" onClick={() => navigate("User")}>
                  <div className="cardContent">
                    <div className="cardalin">
                      <span className="boximage">
                        <img src={UserLogo} width="25" height="25" />
                      </span>
                      <h4>User Dashboard</h4>
                    </div>
                  </div>
                </div>

                <div className="infoCard" onClick={() => navigate("Approver")}>
                  <div className="cardContent">
                    <div className="cardalin">
                      <span className="boximage">
                        <img src={ApproverLogo} width="25" height="25" />
                      </span>
                      <h4>Approver Dashboard</h4>
                    </div>
                  </div>
                </div>

                <div className="infoCard" onClick={() => navigate("Performer")}>
                  <div className="cardContent">
                    <div className="cardalin">
                      <span className="boximage">
                        <img src={ApLogo} width="25" height="25" />
                      </span>
                      <h4>AP Performer Dashboard</h4>
                    </div>
                  </div>
                </div>

              </div>
            </div>
          </section>
        </div>
      )}

      {/* ✅ USER DASHBOARD */}
      {page === "User" && (
        <UserDashboard context={props.context} />
      )}

      {/* ✅ APPROVER DASHBOARD */}
      {page === "Approver" && (
        <ApproverDashboard context={props.context} />
      )}

      {/* ✅ AP PERFORMER DASHBOARD */}
      {page === "Performer" && (
        <APperformerDashboard context={props.context} />
      )}

    </div>
  );
}