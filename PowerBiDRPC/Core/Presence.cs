using PowerBiDRPC.DRPCCore;

using System;

namespace PowerBiDRPC.Core
{

	public class Presence : IDisposable
	{
		private DiscordRpc.EventHandlers handlers = default;
		private DiscordRpc.RichPresence presence;


		/// <summary>
		/// Initializes the Rich Presence to the specified application ID.
		/// </summary>
		/// <param name="appId"></param>
		public void InitializePresence(string appId)
		{
			DiscordRpc.Initialize(appId, ref handlers, true, null);
			presence = new DiscordRpc.RichPresence();
		}


		/// <summary>
		/// Updates the Rich Presence instance.
		/// </summary>
		public void UpdatePresence()
		{
			DiscordRpc.UpdatePresence(ref presence);
		}


		/// <summary>
		/// Updates the Large Image in the Rich Presence instance.
		/// </summary>
		/// <param name="largeImageKey">Large Image Key</param>
		/// <param name="largeImageText">Large Image Text</param>
		public void UpdateLargeImage(string largeImageKey, string largeImageText = "")
		{
			presence.largeImageKey = largeImageKey;
			presence.largeImageText = largeImageText;

			UpdatePresence();
		}


		/// <summary>
		///  Updates the Small Image in the Rich Presence instance.
		/// </summary>
		/// <param name="smallImageKey">Small Image Key</param>
		/// <param name="smallImageText">Small Image Text</param>
		public void UpdateSmallImage(string smallImageKey, string smallImageText = "")
		{
			presence.smallImageKey = smallImageKey;
			presence.smallImageText = smallImageText;

			UpdatePresence();
		}


		/// <summary>
		/// Updates the State in the Rich Presence instance.
		/// </summary>
		/// <param name="state">The new state</param>
		public void UpdateState(string state)
		{
			presence.state = state;

			UpdatePresence();
		}


		/// <summary>
		/// Updates the Details in the Rich Presence instance.
		/// </summary>
		/// <param name="details">The new details</param>
		public void UpdateDetails(string details)
		{
			presence.details = details;

			UpdatePresence();
		}


		/// <summary>
		/// Shuts down the Rich Presence instance.
		/// </summary>
		public void ShutDown()
		{
			DiscordRpc.Shutdown();
		}

		public void Dispose()
		{
			DiscordRpc.Shutdown();

			GC.SuppressFinalize(this);
		}
	}
}