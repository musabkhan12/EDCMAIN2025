import * as React from 'react';
import { useEffect, useState } from 'react';
import { SPFI } from '@pnp/sp';
import "@pnp/sp/site-groups/web";
import "@pnp/sp/profiles";
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { GraphFI, graphfi, SPFx as graphSPFx } from "@pnp/graph";
import "@pnp/graph/groups";
import "@pnp/graph/members";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups/web";
// import "./../../components/DmsMain.scss";
interface IProps {
  sp: SPFI;
  props: any;
}

interface IUser {
  Name: string;
  Email: string;
  Department: string;
}

const IntranetGroupMember: React.FC<IProps> = ({ sp, props }) => {
  const [users, setUsers] = useState<IUser[]>([]);
  const [loading, setLoading] = useState(true);
  let graph: GraphFI;
  useEffect(() => {
    graph = graphfi().using(graphSPFx(props.context));
    loadMembers();
  }, []);

  const loadMembers = async () => {
    try {
      // 1. Get group users
      const groupUsers = await sp.web.siteGroups
        .getByName("Intranet Member Group")
        .users
        .select("*,Title", "Email", "LoginName")();

      // 2. Get Department from User Profile
      console.log("Group Users:", groupUsers);
      // const allusers = await graph.users
      //   .select("*,id,displayName,mail,userPrincipalName,department,Division,Sector")
      //   .top(999)();
      const allusers = await graph.users
        .select(
          "id,displayName,mail,userPrincipalName,department," +
          "onPremisesExtensionAttributes"
        )
        .top(999)();

      console.log(allusers, "allusers");
      const result: IUser[] = await Promise.all(
        groupUsers.map(async (u) => {
          let department = "-";

          try {
            const profile = await sp.profiles.getPropertiesFor(u.LoginName);
            const deptProp = profile.UserProfileProperties?.find(
              (p: any) => p.Key === "Department"
            );
            department = deptProp?.Value || "-";
          } catch (err) {
            console.warn("Profile not found for", u.LoginName);
          }

          return {
            Name: u.Title,
            Email: u.Email,
            Department: department
          };
        })
      );

      setUsers(result);
    } catch (error) {
      console.error("Error loading group members", error);
    } finally {
      setLoading(false);
    }
  };

  const exportToExcel = () => {
    const worksheet = XLSX.utils.json_to_sheet(users);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Members");
    const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });

    saveAs(
      new Blob([excelBuffer], { type: "application/octet-stream" }),
      "IntranetGroupMembers.xlsx"
    );
  };

  return (
    <div>
      <h3>Intranet Group Members</h3>

      <button onClick={exportToExcel} disabled={!users.length}>
        Export to Excel
      </button>

      {loading ? (
        <p>Loading...</p>
      ) : (
        // <table style={{ width: "100%", borderCollapse: "collapse", marginTop: 10 }}>
        //   <thead>
        //     <tr>
        //       <th>Name</th>
        //       <th>Email</th>
        //       <th>Department</th>
        //     </tr>
        //   </thead>
        //   <tbody>
        //     {users.map((u, i) => (
        //       <tr key={i}>
        //         <td>{u.Name}</td>
        //         <td>{u.Email}</td>
        //         <td>{u.Department}</td>
        //       </tr>
        //     ))}
        //   </tbody>
        // </table>
        <table id="tabAllItems" className='mtbalenew'>
          <thead>
            <tr>
              <th style={{ minWidth: '40px', textAlign: 'center', maxWidth: '40px' }}>

                <div style={{
                  width: '100%', height: '80px', clear: 'both', display: 'flex', justifyContent: 'start', textAlign
                    : 'center'
                }} className='pb-3'> Name</div>
              </th>

              <th style={{ minWidth: '50px', textAlign: 'center', maxWidth: '50px' }}>
                <div style={{
                  width: '100%', height: '80px', display: 'flex', justifyContent: 'start', textAlign
                    : 'center'
                }} className='pb-3'>Email</div>

              </th>
              <th style={{ minWidth: '50px', textAlign: 'center', maxWidth: '50px' }}>
                <div style={{
                  width: '100%', height: '80px', display: 'flex', justifyContent: 'start', textAlign
                    : 'center'
                }} className='pb-3'>Department</div>

              </th>
            </tr>
          </thead>


          <tbody>
            {users.map((u, i) => (
              <tr key={i}>
                <td>{u.Name}</td>
                <td>{u.Email}</td>
                <td>{u.Department}</td>
              </tr>
            ))}
          </tbody>


        </table>
      )}
    </div>
  );
};

export default IntranetGroupMember;
