using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace WaiAddIn
{
    public class ChatMessageData
    {
        public string Role { get; set; } = "user";
        public string Content { get; set; } = "";
        public DateTime Timestamp { get; set; } = DateTime.Now;
    }

    public class ChatSessionData
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string Title { get; set; } = "新会话";
        public DateTime CreatedAt { get; set; } = DateTime.Now;
        public DateTime UpdatedAt { get; set; } = DateTime.Now;
        public List<ChatMessageData> Messages { get; set; } = new List<ChatMessageData>();

        public override string ToString()
        {
            return string.IsNullOrWhiteSpace(Title) ? $"会话 {CreatedAt:MM-dd HH:mm}" : Title;
        }
    }

    public static class ConversationStore
    {
        private static readonly string StorePath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "WAI", "sessions.json");

        public static List<ChatSessionData> Load()
        {
            try
            {
                var dir = Path.GetDirectoryName(StorePath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                if (File.Exists(StorePath))
                {
                    var json = File.ReadAllText(StorePath);
                    return JsonConvert.DeserializeObject<List<ChatSessionData>>(json) ?? new List<ChatSessionData>();
                }
            }
            catch { }

            return new List<ChatSessionData>();
        }

        public static void Save(List<ChatSessionData> sessions)
        {
            try
            {
                var dir = Path.GetDirectoryName(StorePath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                var json = JsonConvert.SerializeObject(sessions ?? new List<ChatSessionData>(), Formatting.Indented);
                File.WriteAllText(StorePath, json);
            }
            catch { }
        }

        public static ChatSessionData CreateSession(string title = "新会话")
        {
            return new ChatSessionData
            {
                Id = Guid.NewGuid().ToString("N"),
                Title = title,
                CreatedAt = DateTime.Now,
                UpdatedAt = DateTime.Now,
                Messages = new List<ChatMessageData>()
            };
        }

        public static ChatSessionData GetOrCreateById(List<ChatSessionData> sessions, string sessionId)
        {
            if (sessions == null)
            {
                sessions = new List<ChatSessionData>();
            }

            var session = sessions.FirstOrDefault(s => s.Id == sessionId);
            if (session != null)
            {
                return session;
            }

            session = CreateSession();
            sessions.Insert(0, session);
            return session;
        }
    }
}
