using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel;

namespace cscmdlets
{
    public class CSClients
    {

        private String servicesDirectory;

        public Authentication.AuthenticationClient authClient;
        public Authentication.OTAuthentication authAuth;
        private DateTime? sessionExpiry;
        private Boolean authClientOpen = false;
        private String AuthenticationEndpointAddress = "Authentication.svc";

        // todo later support other authentication methods
        private String AuthenticationMethod;
        private String userName;
        private String password;

        public DocumentManagement.DocumentManagementClient docClient;
        public DocumentManagement.OTAuthentication docAuth;
        private Boolean docClientOpen = false;
        private String DocumentManagementEndpointAddress = "DocumentManagement.svc";

        public MemberService.MemberServiceClient memberClient;
        public MemberService.OTAuthentication memberAuth;
        private Boolean memberClientOpen = false;
        private String MemberServiceEndpointAddress = "MemberService.svc";

        public Collaboration.CollaborationClient collabClient;
        public Collaboration.OTAuthentication collabAuth;
        private Boolean collabClientOpen = false;
        private String CollaborationEndpointAddress = "Collaboration.svc";

        public Classifications.ClassificationsClient classClient;
        public Classifications.OTAuthentication classAuth;
        private Boolean classClientOpen = false;
        private String ClassificationsEndpointAddress = "Classifications.svc";

        public RecordsManagement.RecordsManagementClient rmClient;
        public RecordsManagement.OTAuthentication rmAuth;
        private Boolean rmClientOpen = false;
        private String RecordsManagementEndpointAddress = "RecordsManagement.svc";

        public PhysicalObjects.PhysicalObjectsClient poClient;
        public PhysicalObjects.OTAuthentication poAuth;
        private Boolean poClientOpen = false;
        private String PhysicalObjectsEndpointAddress = "PhysicalObjects.svc";
        
        public CSClients(String ServicesDirectory)
        {
            if (!ServicesDirectory.EndsWith("/"))
                servicesDirectory = ServicesDirectory + "/";
            else
                servicesDirectory = ServicesDirectory;

            AuthenticationEndpointAddress = servicesDirectory + AuthenticationEndpointAddress;
            CollaborationEndpointAddress = servicesDirectory + CollaborationEndpointAddress;
            DocumentManagementEndpointAddress = servicesDirectory + DocumentManagementEndpointAddress;
            MemberServiceEndpointAddress = servicesDirectory + MemberServiceEndpointAddress;
            ClassificationsEndpointAddress = servicesDirectory + ClassificationsEndpointAddress;
            RecordsManagementEndpointAddress = servicesDirectory + RecordsManagementEndpointAddress;
            PhysicalObjectsEndpointAddress = servicesDirectory + PhysicalObjectsEndpointAddress;
        }

        public void AddAuthenticationDetails(String UserName, String Password)
        {
            AuthenticationMethod = "user";
            userName = UserName;
            if (Password.StartsWith("!=!enc!=!"))
                password = Encryption.DecryptString(Password);
            else
                password = Password;
        }

        public void OpenClient(Type ClientType)
        {
            try
            {

                // make the initial connection if it hasn't already been made
                if (!authClientOpen)
                    GetToken();

                // create the generic binding object
                BasicHttpBinding binding = new BasicHttpBinding();
                binding.SendTimeout = new TimeSpan(0, 0, 0, 0, 100000);
                binding.OpenTimeout = new TimeSpan(0, 0, 0, 0, 100000);

                // create the client
                if (ClientType.Equals(typeof(Authentication.AuthenticationClient)))
                {
                    // it's already open, just catching this to stop throwing an error
                    return;
                }
                else if (ClientType.Equals(typeof(DocumentManagement.DocumentManagementClient)))
                {
                    // open the client
                    EndpointAddress address = new EndpointAddress(DocumentManagementEndpointAddress);
                    docClient = new DocumentManagement.DocumentManagementClient(binding, address);
                    docClient.Open();
                    docClientOpen = true;

                    // create the authentication object
                    docAuth = new DocumentManagement.OTAuthentication();
                    docAuth.AuthenticationToken = authAuth.AuthenticationToken;
                    return;
                }
                else if (ClientType.Equals(typeof(MemberService.MemberServiceClient)))
                {
                    // open the client
                    EndpointAddress address = new EndpointAddress(MemberServiceEndpointAddress);
                    memberClient = new MemberService.MemberServiceClient(binding, address);
                    memberClient.Open();
                    memberClientOpen = true;

                    // create the authentication object
                    memberAuth = new MemberService.OTAuthentication();
                    memberAuth.AuthenticationToken = authAuth.AuthenticationToken;
                    return;
                }
                else if (ClientType.Equals(typeof(Collaboration.CollaborationClient)))
                {
                    // open the client
                    EndpointAddress address = new EndpointAddress(CollaborationEndpointAddress);
                    collabClient = new Collaboration.CollaborationClient(binding, address);
                    collabClient.Open();
                    collabClientOpen = true;

                    // create the authentication object
                    collabAuth = new Collaboration.OTAuthentication();
                    collabAuth.AuthenticationToken = authAuth.AuthenticationToken;
                    return;
                }
                else if (ClientType.Equals(typeof(Classifications.ClassificationsClient)))
                {
                    // open the client
                    EndpointAddress address = new EndpointAddress(ClassificationsEndpointAddress);
                    classClient = new Classifications.ClassificationsClient(binding, address);
                    classClient.Open();
                    classClientOpen = true;

                    // create the authentication object
                    classAuth = new Classifications.OTAuthentication();
                    classAuth.AuthenticationToken = authAuth.AuthenticationToken;
                    return;
                }
                else if (ClientType.Equals(typeof(RecordsManagement.RecordsManagementClient)))
                {
                    // open the client
                    EndpointAddress address = new EndpointAddress(RecordsManagementEndpointAddress);
                    rmClient = new RecordsManagement.RecordsManagementClient(binding, address);
                    rmClient.Open();
                    rmClientOpen = true;

                    // create the authentication object
                    rmAuth = new RecordsManagement.OTAuthentication();
                    rmAuth.AuthenticationToken = authAuth.AuthenticationToken;
                    return;
                }
                else if (ClientType.Equals(typeof(PhysicalObjects.PhysicalObjectsClient)))
                {
                    // open the client
                    EndpointAddress address = new EndpointAddress(PhysicalObjectsEndpointAddress);
                    poClient = new PhysicalObjects.PhysicalObjectsClient(binding, address);
                    poClient.Open();
                    poClientOpen = true;

                    // create the authentication object
                    poAuth = new PhysicalObjects.OTAuthentication();
                    poAuth.AuthenticationToken = authAuth.AuthenticationToken;
                    return;
                }

                throw new Exception("Client type not supported");

            }
            catch (Exception e)
            {
                throw e;
            }
        }
        
        public void CheckSession()
        {
            // reset the tokens if the session has expired
            if (authClient == null)
                throw new Exception("Session not established.");
            else if (sessionExpiry != null && sessionExpiry < DateTime.Now)
            {
                // todo later support other authentication methods

                // get the new token and reset expiry
                String token = authClient.AuthenticateUser(userName, password);
                authAuth.AuthenticationToken = token;
                sessionExpiry = authClient.GetSessionExpirationDate(ref authAuth);

                // reset any tokens on the client authentication objects
                if (docAuth != null) docAuth.AuthenticationToken = token;
                if (memberAuth != null) memberAuth.AuthenticationToken = token;
                if (collabAuth != null) collabAuth.AuthenticationToken = token;
                if (classAuth != null) classAuth.AuthenticationToken = token;
                if (rmAuth != null) rmAuth.AuthenticationToken = token;
                if (poAuth != null) poAuth.AuthenticationToken = token;
            }
        }

        private void GetToken()
        {
            try
            {
                // todo later support other authentication methods

                // check we've got an authentication method
                if (String.IsNullOrEmpty(AuthenticationMethod))
                    throw new Exception("No authentication detasil supplied");
                
                // create the generic binding object
                BasicHttpBinding binding = new BasicHttpBinding();
                binding.SendTimeout = new TimeSpan(0, 0, 0, 0, 100000);
                binding.OpenTimeout = new TimeSpan(0, 0, 0, 0, 100000);

                // open the authentication client
                EndpointAddress address = new EndpointAddress(AuthenticationEndpointAddress);
                authClient = new Authentication.AuthenticationClient(binding, address);
                authClient.Open();
                authClientOpen = true;
            
                // get the authentication token
                String token = authClient.AuthenticateUser(userName, password);

                // create the authentication object
                authAuth = new Authentication.OTAuthentication();
                authAuth.AuthenticationToken = token;

                // get the session expiry
                sessionExpiry = authClient.GetSessionExpirationDate(ref authAuth);

            }
            catch (Exception e)
            {
                throw e;
            }

        }

        internal void CloseClients()
        {
            if (authClientOpen) authClient.Close();
            if (collabClientOpen) collabClient.Close();
            if (docClientOpen) docClient.Close();
            if (memberClientOpen) memberClient.Close();
            if (classClientOpen) classClient.Close();
            if (rmClientOpen) rmClient.Close();
            if (poClientOpen) poClient.Close();
        }

    }
}
