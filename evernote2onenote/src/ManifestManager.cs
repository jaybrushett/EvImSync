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
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Evernote2Onenote
{
    internal static class ManifestManager
    {
        private static readonly object _fileLock = new object();

        /// <summary>
        /// Creates a new manifest file in logFolder for the given jobs and returns its path.
        /// </summary>
        public static string Create(string logFolder, IEnumerable<ImportJob> jobs)
        {
            Directory.CreateDirectory(logFolder);
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var path = Path.Combine(logFolder, $"e2o_import_manifest_{timestamp}.txt");
            var lines = jobs.Select(j => new ManifestEntry
            {
                RelativePath  = j.RelativePath,
                OneNoteTarget = j.OneNoteTargetPath,
                AttemptCount  = 0,
                Status        = ManifestStatus.Pending
            }.Serialize());
            File.WriteAllLines(path, lines, Encoding.UTF8);
            return path;
        }

        /// <summary>
        /// Loads all valid entries from a manifest file.
        /// </summary>
        public static List<ManifestEntry> Load(string filePath)
        {
            if (!File.Exists(filePath))
                return new List<ManifestEntry>();

            lock (_fileLock)
            {
                return File.ReadAllLines(filePath, Encoding.UTF8)
                    .Select(ManifestEntry.TryParse)
                    .Where(e => e != null)
                    .ToList();
            }
        }

        /// <summary>
        /// Replaces the matching entry (by RelativePath) in the manifest file.
        /// </summary>
        public static void UpdateEntry(string filePath, ManifestEntry updated)
        {
            lock (_fileLock)
            {
                var lines = File.ReadAllLines(filePath, Encoding.UTF8);
                for (int i = 0; i < lines.Length; i++)
                {
                    var entry = ManifestEntry.TryParse(lines[i]);
                    if (entry != null && entry.RelativePath == updated.RelativePath)
                    {
                        lines[i] = updated.Serialize();
                        break;
                    }
                }
                File.WriteAllLines(filePath, lines, Encoding.UTF8);
            }
        }

        /// <summary>
        /// Checks the most recent manifest file in logFolder for resumable work.
        /// Returns the file path if found, null if nothing to resume.
        /// </summary>
        public static string FindResumable(string logFolder)
        {
            if (!Directory.Exists(logFolder))
                return null;

            var newest = Directory.GetFiles(logFolder, "e2o_import_manifest_*.txt")
                .OrderByDescending(f => f)
                .FirstOrDefault();

            if (newest == null)
                return null;

            var entries = Load(newest);
            bool hasWork = entries.Any(e =>
                e.Status == ManifestStatus.Pending ||
                e.Status == ManifestStatus.InProgress ||
                (e.Status == ManifestStatus.Failed && e.AttemptCount < 3));

            return hasWork ? newest : null;
        }
    }
}
