import React, { useState, useEffect } from 'react';
import { PublicClientApplication } from '@azure/msal-browser';
import Cookies from 'js-cookie';
import './Button.css';

const msalConfig = {
  auth: {
    clientId: '99f25316-47bf-4c6d-b0ac-4b178de98c42',
    authority: 'https://login.microsoftonline.com/3c90a2ff-691c-483a-8e94-fbca1b7d4edf'
    
  },
};

const pca = new PublicClientApplication(msalConfig);

const getAccessToken = async () => {
  const request = {
    scopes: ["User.Read", "Sites.ReadWrite.All", "Sites.Manage.All"]
  };

  try {
    let accessToken = Cookies.get('accessToken');
    console.log('Cached accessToken:', accessToken);

    if (!accessToken) {
      await pca.initialize();
      const loginResponse = await pca.loginPopup(request);
      console.log('Login response:', loginResponse);
      accessToken = loginResponse.accessToken;

      // Set cookie with an expiration time (e.g., 1 hour)
      const expirationTime = new Date(response.expiresOn); // 1 hour
      Cookies.set('accessToken', accessToken, { expires: expirationTime });
    }
    return accessToken;
  } catch (error) {
    console.error('Error getting access token:', error);
    throw error; // Re-throw to handle in the calling function
  }
}

const GraphAPI = (props) => {
  const [data, setData] = useState([]);

  const getList = async () => {
    console.log("You have clicked login");

    try {
      const accessToken = await getAccessToken();
      console.log('Access token:', accessToken);

      const myHeaders = new Headers();
      myHeaders.append("Authorization", `Bearer ${accessToken}`);

      const requestOptions = {
        method: 'GET',
        headers: myHeaders,
        redirect: 'follow'
      };

      const graphResponse = await fetch("https://graph.microsoft.com/v1.0/sites/xrmlabs.sharepoint.com,e5380930-eb71-4161-b333-20dc135c9109,82f02cd9-b117-4525-a3a1-4db73b525b84/lists/f5aad320-b91b-441e-a2a9-7cd0545aeffe/items?expand=fields", requestOptions);

      if (!graphResponse.ok) {
        throw new Error(`HTTP error! Status: ${graphResponse.status}`);
      }

      const result = await graphResponse.json();
      console.log("Graph API Response:", result);
      console.log(result.value);
      setData(result.value);

    } catch (error) {
      console.error('Error authenticating or fetching data:', error);
    }
  }
  const createList = async () => {
    console.log("You have clicked login");

    try {
      const accessToken = await getAccessToken();
      console.log('Access token:', accessToken);

      const myHeaders = new Headers();
      myHeaders.append("Authorization", `Bearer ${accessToken}`);
      myHeaders.append("Content-Type", "application/json");
      const listBody = JSON.stringify({
        "displayName": "SPFXGraphAPI",
        "columns": [
            {
                "name": "Author",
                "text": {}
            },
            {
                "name": "PageCount",
                "number": {}
            }
        ],
        "list": {
            "template": "genericList"
        }
    });

      const requestOptions = {
        method: 'POST',
        headers: myHeaders,
        body: listBody,
        redirect: 'follow'
      };

      const graphResponse = await fetch("https://graph.microsoft.com/v1.0/sites/xrmlabs.sharepoint.com,e5380930-eb71-4161-b333-20dc135c9109,82f02cd9-b117-4525-a3a1-4db73b525b84/lists", requestOptions);

      if (!graphResponse.ok) {
        throw new Error(`HTTP error! Status: ${graphResponse.status}`);
      }

      const result = await graphResponse.json();
      console.log("Graph API Response:", result);
      console.log(result.value);
      setData(result.value);

    } catch (error) {
      console.error('Error authenticating or fetching data:', error);
    }
  }
    const create = async ()=>{
    console.log('you have clicked create');
    try {

      let accessToken = await getAccessToken();
      console.log(accessToken);

      const myHeaders = new Headers();
      myHeaders.append("Authorization", `Bearer ${accessToken}`);
      myHeaders.append("Content-Type", "application/json");

      const requestOptions = {
        method: 'POST',
        headers: myHeaders,
        body: JSON.stringify({
          fields: {
       
            Title: "New Onboarding",
            Dateofevent: "2024-08-06T18:30:00Z"
            
            
          }
        })
      };

      const graphresponse = await fetch('https://graph.microsoft.com/v1.0/sites/xrmlabs.sharepoint.com,e5380930-eb71-4161-b333-20dc135c9109,82f02cd9-b117-4525-a3a1-4db73b525b84/lists/f5aad320-b91b-441e-a2a9-7cd0545aeffe/items',requestOptions)
      if(graphresponse.ok){
        alert(`Item Created with ID ${graphresponse.id}`)
      }

    } catch (error) {
      console.log(error)
      
    }
  }
  const update = async (ID)=>{
    console.log('you have clicked create');
    try {

      let accessToken = await getAccessToken();
      console.log(accessToken);

      const myHeaders = new Headers();
      myHeaders.append("Authorization", `Bearer ${accessToken}`);
      myHeaders.append("Content-Type", "application/json");

      const requestOptions = {
        method: 'PATCH',
        headers: myHeaders,
        body: JSON.stringify({
          
       
            Title: "SharePoint KT Session",
           
          
        })
      };

      const graphresponse = await fetch(`https://graph.microsoft.com/v1.0/sites/xrmlabs.sharepoint.com,e5380930-eb71-4161-b333-20dc135c9109,82f02cd9-b117-4525-a3a1-4db73b525b84/lists/f5aad320-b91b-441e-a2a9-7cd0545aeffe/items/${ID}/fields`,requestOptions)
      if(graphresponse.ok){
        alert(`Item Updated with ID ${ID}`)
      }

    } catch (error) {
      
    }
  }
  const Delete = async (ID) => {
    let accessToken = await getAccessToken();
    if (!accessToken) {
      console.error("Access token is missing.");
      return;
    }
 
    try {
      const myHeaders = new Headers();
      myHeaders.append("Authorization", `Bearer ${accessToken}`);
 
      const requestOptions = {
        method: 'DELETE',
        headers: myHeaders,
        redirect: 'follow'
      };
 
      const graphResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/xrmlabs.sharepoint.com,e5380930-eb71-4161-b333-20dc135c9109,82f02cd9-b117-4525-a3a1-4db73b525b84/lists/f5aad320-b91b-441e-a2a9-7cd0545aeffe/items/${ID}`, requestOptions);
 
      if (graphResponse.ok) {
        alert(`Item with ID ${ID} is deleted successfully`);
      } else {
        const errorDetails = await graphResponse.json();
        console.error("Error deleting item:", errorDetails);
        alert(`Failed to delete item with ID ${ID}. Status: ${graphResponse.status}`);
      }
    } catch (error) {
      console.error("Error:", error);
    }
  }

  return (
    <>
    <div>
      <h1>Microsoft GraphAPI</h1>
      <div>
        <button className='Button' onClick={getList}>Get Item</button>&nbsp;&nbsp;
        <button className='Button1' onClick={createList}>Create list</button>&nbsp;&nbsp;
        <button className='Button2'>Get Item</button>&nbsp;&nbsp;
        <button className='Button3' onClick={()=>create()}>Create Item</button>&nbsp;&nbsp;
        <button className='Button4' onClick={()=>update(3)}>Update Item</button>&nbsp;&nbsp;
        <button className='Button5' onClick={()=>Delete(2)}>Delete Item</button>&nbsp;&nbsp;
      </div>
    </div>
    <h1>Celebration Data Table</h1>
    {data.length > 0 && (
    <div>
    <table>
         <tr>
            <th>Celebration</th>
            <th>Date of Event</th>
         </tr>
         {data.map(item => (
         <tr key={item.id}>
            <td>{item.fields.Title}</td>
            <td>{item.fields.Dateofevent}</td>
        </tr>
         ))}
    </table>
    </div>
    )}
    </>
  )
}

export default GraphAPI;
