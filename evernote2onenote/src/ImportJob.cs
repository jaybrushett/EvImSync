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

namespace Evernote2Onenote
{
    /// <summary>
    /// Describes one .enex file and its target position in the OneNote hierarchy.
    /// </summary>
    internal class ImportJob
    {
        /// <summary>Absolute path to the .enex file on disk.</summary>
        public string EnexPath { get; set; }

        /// <summary>Path relative to the root import folder (used as the manifest key).</summary>
        public string RelativePath { get; set; }

        /// <summary>Section group name under the notebook, or null for a direct-child section.</summary>
        public string SectionGroupName { get; set; }

        /// <summary>OneNote section name (derived from the ENEX filename without extension).</summary>
        public string SectionName { get; set; }

        /// <summary>Human-readable target path written into the manifest, e.g. "Notebook\Group\Section".</summary>
        public string OneNoteTargetPath { get; set; }
    }
}
