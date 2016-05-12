using Microsoft.AspNet.SignalR;
using Microsoft.AspNet.SignalR.Hubs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConnectingtoO365APIWeb.Hubs
{
    [HubName("oauth")]
    public class OAuthHub : Hub
    {
        public override Task OnConnected()
        {
            Groups.Add(Context.ConnectionId, Context.ConnectionId);
            Clients.Group(Context.ConnectionId).connected(Context.ConnectionId);
            return base.OnConnected();
        }

        public override Task OnDisconnected(bool stopCalled)
        {
            Groups.Remove(Context.ConnectionId, Context.ConnectionId);
            return base.OnDisconnected(stopCalled);
        }
    }
}
