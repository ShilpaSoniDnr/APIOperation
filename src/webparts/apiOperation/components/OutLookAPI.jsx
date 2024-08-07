import React from 'react'
import React, { useState } from 'react';
import { PublicClientApplication } from '@azure/msal-browser';
import Cookies from 'js-cookie';
import './Button.css';

const getAccessToken = async () => {
    const msalConfig = {
      auth: {
        clientId: '99f25316-47bf-4c6d-b0ac-4b178de98c42',
        authority: 'https://login.microsoftonline.com/3c90a2ff-691c-483a-8e94-fbca1b7d4edf',
      },
    };
    const pca = new PublicClientApplication(msalConfig);
  
    const request = {
      scopes: ["User.Read", "Calendars.Read","Sites.ReadWrite.All","Mail.Read"]
    };
  
    try {
      let accessToken = Cookies.get('accessToken');
      console.log(accessToken);
  
      if (!accessToken) {
        await pca.initialize();
        const response = await pca.loginPopup(request);
        console.log(response);
        accessToken = response.accessToken;
        const expirationTime = new Date(response.expiresOn);
        Cookies.set('accessToken', accessToken, { expires: expirationTime });
      }
      return accessToken;
    } catch (error) {
      console.error('Error getting access token:', error);
      throw error; // Re-throw to handle in the calling function
    }
  }

const OutLookAPI = (props) => {
    const [events, setEvents] = React.useState([]);
    const [mails, setMails] = React.useState([]);

    const getEvents = async () => {
        console.log("You have clicked get events");
    
        try {
          let accessToken = await getAccessToken();
          console.log(accessToken);
    
          const myHeaders = new Headers();
          myHeaders.append("Authorization", `Bearer ${accessToken}`);
    
          const requestOptions = {
            method: 'GET',
            headers: myHeaders,
            redirect: 'follow'
          };
    
          const graphResponse = await fetch("https://graph.microsoft.com/v1.0/me/events", requestOptions);
    
          if (!graphResponse.ok) {
            throw new Error(`HTTP error! Status: ${graphResponse.status}`);
          }
    
          const result = await graphResponse.json();
          console.log("Graph API Response:", result);
          setEvents(result.value);
    
        } catch (error) {
          console.error('Error authenticating or fetching data:', error);
        }
      };

    const getEmails = async ()=>{
        const accessToken = await getAccessToken();
        try {
          
          const myHeaders = new Headers();
          myHeaders.append("Authorization", `Bearer ${accessToken}`);
    
          const requestOptions = {
            method: 'GET',
            headers: myHeaders,
            redirect: 'follow'
          };
    
          const response = await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders('Inbox')/messages`,requestOptions);
          const result = await response.json();
          
          console.log(result.value);
          setMails(result.value)
          
    
    
    
    
        } catch (error) {
          console.log(error)
        }
      }
  return (
    <>
    <div>
    <h1>OutLookAPI</h1>
    <div>
    <button className='Button' onClick={getEvents}>Get Events</button>&nbsp;&nbsp;
    <button className='Button1' onClick={getEmails}>Get Mails</button>&nbsp;&nbsp;
    </div>
    </div>
    <h1>Outlook Event Table</h1>
    {events.length > 0 ? (
    <div >
    <table>
         <tr>
            <th>Organizer</th>
            <th>Subject</th>
            <th>Start</th>
            <th>End</th>
            <th>Location</th>
         </tr>
         {events.map(event => (
         <tr key={event.id}>
            <td>{event.organizer.emailAddress.name}</td>
            <td>{event.subject}</td>
            <td>{new Date(event.start.dateTime).toLocaleString()}</td>
            <td>{new Date(event.end.dateTime).toLocaleString()}</td>
            <td>{event.location.displayName || 'N/A'}</td>
        </tr>
        
         ))}
    </table>
    </div>
    ) : (
        <div>
            <h1>No events found</h1>
          </div>
        )
    }
    {mails.length > 0 ? (
   
    <div>
    <table>
         <tr>
            <th>Sender</th>
            <th>Subject</th>
            <th>From</th>
            <th>To Receipient</th>
            <th>Importance</th>
         </tr>
         { mails.map(mail => (
         <tr key={mail.id}>
            <td>{mail.sender.emailAddress.name}</td>
            <td>{mail.subject}</td>
            <td>{mail.from.emailAddress.address}</td>
            <td>{mail.toRecipients[0].emailAddress.address}</td>
            <td>{mail.importance}</td>
        </tr>
        
         ))}
    </table>
    </div>
    ) : (
        <div>
            <h1>No mails found</h1>
          </div>
        )
    }
    </>
    
  )
}


export default OutLookAPI