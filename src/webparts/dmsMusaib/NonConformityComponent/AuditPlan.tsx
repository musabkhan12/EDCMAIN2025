import * as React from "react";
import styles from "./AuditPlan.module.scss";
import type { IAuditPlanProps } from "./IAuditPlanProps";
import { HashRouter as Router, Routes, Route } from "react-router-dom";
import ListingPage from "./ListingForm";
import HomePage from "./AddForm";
import EditFormWrapper from "./EditFormWrapper";

export default class AuditPlan extends React.Component<IAuditPlanProps, {}> {
  public render(): React.ReactElement<IAuditPlanProps> {
    return (
      <section className={styles.auditPlan}>
        <div className={styles.welcome}>
          <Router>
            <Routes>
              <Route
                path="/listing"
                element={
                  <ListingPage
                    description=""
                    context={this.props.context}
                    currentUserID={this.props.currentUserID}
                    userDisplayName={this.props.userDisplayName}
                  />
                }
              />
              <Route
                // path="/new"
                path="/NonConformity"
                element={
                  <HomePage
                    description=""
                    context={this.props.context}
                    currentUserID={this.props.currentUserID}
                    userDisplayName={this.props.userDisplayName}
                  />
                }
              />
              <Route
                path="/:type/:postId"
                element={
                  <EditFormWrapper
                    description=""
                    context={this.props.context}
                    currentUserID={this.props.currentUserID}
                    userDisplayName={this.props.userDisplayName}
                  />
                }
              />
              <Route
                path="/:type/:postId/:itmId"
                element={
                  <EditFormWrapper
                    description=""
                    context={this.props.context}
                    currentUserID={this.props.currentUserID}
                    userDisplayName={this.props.userDisplayName}
                  />
                }
              />
            </Routes>
          </Router>
        </div>
      </section>
    );
  }
}
