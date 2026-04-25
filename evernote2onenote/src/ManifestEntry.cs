// Evernote2Onenote - imports Evernote notes to Onenote
// Copyright (C) 2014, 2023 - Stefan Kueng

// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.

// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.

using System;

namespace Evernote2Onenote
{
    internal enum ManifestStatus
    {
        Pending,
        InProgress,
        Complete,
        Failed
    }

    /// <summary>
    /// Represents one line in the import manifest file.
    /// Serialized format: [STATUS] relativePath | oneNoteTarget | attemptCount
    /// </summary>
    internal class ManifestEntry
    {
        public string RelativePath { get; set; }
        public string OneNoteTarget { get; set; }
        public int AttemptCount { get; set; }
        public ManifestStatus Status { get; set; }

        public string Serialize() =>
            $"[{Status.ToString().ToUpper()}] {RelativePath} | {OneNoteTarget} | {AttemptCount}";

        public static ManifestEntry TryParse(string line)
        {
            if (string.IsNullOrWhiteSpace(line))
                return null;

            var statusEnd = line.IndexOf(']');
            if (statusEnd < 2 || line[0] != '[')
                return null;

            var statusStr = line.Substring(1, statusEnd - 1);
            var rest = line.Substring(statusEnd + 1).TrimStart();

            var parts = rest.Split(new[] { " | " }, StringSplitOptions.None);
            if (parts.Length != 3)
                return null;

            if (!Enum.TryParse(statusStr, true, out ManifestStatus status))
                return null;

            if (!int.TryParse(parts[2].Trim(), out int attempts))
                return null;

            return new ManifestEntry
            {
                Status       = status,
                RelativePath = parts[0].Trim(),
                OneNoteTarget = parts[1].Trim(),
                AttemptCount = attempts
            };
        }
    }
}
