<#
    Reads and writes Audacity project files directly.

    Audacity 4.0 dropped mod-script-pipe - there is no scripting interface in it
    at all - so lib\AudacityPipe.ps1 cannot reach a 4.0 project. What 4.0 did
    keep is the file format: .aup4 is SQLite with the same project and autosave
    tables as .aup3, holding the same ProjectSerializer binary XML. That is the
    way in, and because the format did not change it is one way in for both
    versions.

    Dot-source this file; it defines functions only and runs nothing.

    The binary layer is C# rather than PowerShell: it is byte work over
    documents that reach hundreds of kilobytes on a long project, and 5.1 is too
    slow at that to keep a hotkey feeling like a hotkey.

    ---- the format, as verified against 3.7.9 and 4.0 files on 2026-09-09 ----

    Two BLOBs per row. "dict" maps 16-bit ids to names, "doc" is the document
    referring to them. Both are streams of [type byte][payload]:

        0  CharSize   1 byte              8  SizeT     id + 4 bytes  (see below)
        1  StartTag   id                  9  Float     id + 4 + 4 digits
        2  EndTag     id                 10  Double    id + 8 + 4 digits
        3  String     id + 4-byte len    11  Data      4-byte len + text
        4  Int        id + 4             12  Raw       4-byte len + text
        5  Bool       id + 1             13  Push      -
        6  Long       id + 4             14  Pop       -
        7  LongLong   id + 8             15  Name      id + 2-byte len + text

    Lengths are byte counts and text is UTF-16LE. SizeT is written as FOUR
    bytes, not the eight sizeof(size_t) implies on a 64-bit build - the parser
    below assumes four and re-parses with eight if that fails to consume the
    document, so a future build that changes its mind is handled rather than
    silently misread.

    Label tracks look like this, and this is exactly what the writer emits:

        <labeltrack name=String isSelected=Bool height=Int minimized=Bool
                    numlabels=Int>
          <label t=Double t1=Double title=String/>

    Audacity also writes selLow/selHigh on a label, carrying whatever spectral
    selection happened to be live when it was made. They are optional on read
    and mean "this label covers these frequencies", which is not true of a label
    built from a workbook, so they are left off rather than invented.
#>

if (-not ([System.Management.Automation.PSTypeName]'AutoNyx.AuProject').Type) {
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace AutoNyx
{
    public class AuLabel
    {
        public double T;
        public double T1;
        public string Title;
    }

    public class AuLabelTrack
    {
        public string Name;
        public List<AuLabel> Labels = new List<AuLabel>();
        public int DocStart = -1;   // offset of its StartTag byte
        public int DocEnd = -1;     // offset just past its EndTag
    }

    // One project document: the two blobs, plus what we understand of them.
    public class AuProject
    {
        const byte FT_CharSize = 0,  FT_StartTag = 1,  FT_EndTag = 2,  FT_String   = 3,
                   FT_Int      = 4,  FT_Bool     = 5,  FT_Long   = 6,  FT_LongLong = 7,
                   FT_SizeT    = 8,  FT_Float    = 9,  FT_Double = 10, FT_Data     = 11,
                   FT_Raw      = 12, FT_Push     = 13, FT_Pop    = 14, FT_Name     = 15;

        public byte[] Dict;
        public byte[] Doc;
        public string Table;              // which table the blobs came from
        public string Path;
        public long RowId;
        public string AudacityVersion;
        public int SizeTWidth = 4;
        public List<AuLabelTrack> LabelTracks = new List<AuLabelTrack>();
        public List<string> TrackNames = new List<string>();
        public int ProjectEndOffset = -1; // offset of </project>: where tracks are appended

        Dictionary<int, string> names = new Dictionary<int, string>();
        Dictionary<string, int> ids = new Dictionary<string, int>(StringComparer.Ordinal);
        int maxId = -1;

        AuLabelTrack curTrack;
        AuLabel curLabel;

        static ushort ReadU16(byte[] b, ref int i) { ushort v = (ushort)(b[i] | (b[i + 1] << 8)); i += 2; return v; }
        static uint ReadU32(byte[] b, ref int i) { uint v = BitConverter.ToUInt32(b, i); i += 4; return v; }
        string NameOf(int id) { string s; return names.TryGetValue(id, out s) ? s : ("?" + id); }

        public static AuProject Parse(byte[] dict, byte[] doc, string table, string path, long rowId)
        {
            Exception first = null;
            foreach (int width in new int[] { 4, 8 })
            {
                AuProject p = new AuProject();
                p.Dict = dict; p.Doc = doc; p.Table = table; p.Path = path; p.RowId = rowId;
                p.SizeTWidth = width;
                try
                {
                    if (dict != null) p.Scan(dict, true);
                    p.Scan(doc, false);
                    if (p.ProjectEndOffset < 0)
                        throw new Exception("no closing project tag - this is not an Audacity project document");
                    return p;
                }
                catch (Exception ex) { if (first == null) first = ex; }
            }
            throw new Exception("Could not read the project document: " + first.Message);
        }

        void Scan(byte[] buf, bool isDict)
        {
            int i = 0, n = buf.Length;
            Stack<string> stack = new Stack<string>();
            if (!isDict) { curTrack = null; curLabel = null; }

            while (i < n)
            {
                int tagStart = i;
                byte t = buf[i]; i++;
                switch (t)
                {
                    case FT_CharSize: i += 1; break;
                    case FT_Push: case FT_Pop: break;

                    case FT_Name:
                    {
                        int id = ReadU16(buf, ref i);
                        int len = ReadU16(buf, ref i);
                        string nm = Encoding.Unicode.GetString(buf, i, len); i += len;
                        names[id] = nm; ids[nm] = id;
                        if (id > maxId) maxId = id;
                        break;
                    }
                    case FT_StartTag:
                    {
                        string nm = NameOf(ReadU16(buf, ref i));
                        stack.Push(nm);
                        if (!isDict)
                        {
                            if (nm == "labeltrack") { curTrack = new AuLabelTrack(); curTrack.DocStart = tagStart; }
                            else if (nm == "label" && curTrack != null) curLabel = new AuLabel();
                        }
                        break;
                    }
                    case FT_EndTag:
                    {
                        string nm = NameOf(ReadU16(buf, ref i));
                        if (stack.Count > 0) stack.Pop();
                        if (!isDict)
                        {
                            if (nm == "label" && curTrack != null && curLabel != null)
                            { curTrack.Labels.Add(curLabel); curLabel = null; }
                            else if (nm == "labeltrack" && curTrack != null)
                            { curTrack.DocEnd = i; LabelTracks.Add(curTrack); curTrack = null; }
                            else if (nm == "project") ProjectEndOffset = tagStart;
                        }
                        break;
                    }
                    case FT_String:
                    {
                        string an = NameOf(ReadU16(buf, ref i));
                        int len = (int)ReadU32(buf, ref i);
                        string v = Encoding.Unicode.GetString(buf, i, len); i += len;
                        if (!isDict) SetAttr(stack.Count > 0 ? stack.Peek() : null, an, v, 0);
                        break;
                    }
                    case FT_Double:
                    {
                        string an = NameOf(ReadU16(buf, ref i));
                        double v = BitConverter.ToDouble(buf, i); i += 8;
                        i += 4;                                   // digits
                        if (!isDict) SetAttr(stack.Count > 0 ? stack.Peek() : null, an, null, v);
                        break;
                    }
                    case FT_Float:    i += 2; i += 8; break;      // value + digits
                    case FT_Int: case FT_Long: i += 2; i += 4; break;
                    case FT_Bool:     i += 2; i += 1; break;
                    case FT_LongLong: i += 2; i += 8; break;
                    case FT_SizeT:    i += 2; i += SizeTWidth; break;
                    case FT_Data: case FT_Raw:
                    {
                        int len = (int)ReadU32(buf, ref i); i += len; break;
                    }
                    default:
                        throw new Exception("unknown field type " + t + " at offset " + tagStart);
                }
                if (i > n) throw new Exception("field at offset " + tagStart + " runs past the end of the document");
            }
        }

        void SetAttr(string elem, string attr, string s, double d)
        {
            if (elem == null) return;
            if (elem == "project" && attr == "audacityversion") AudacityVersion = s;
            else if (elem == "labeltrack" && curTrack != null && attr == "name") curTrack.Name = s;
            else if (elem == "label" && curLabel != null)
            {
                if (attr == "t") curLabel.T = d;
                else if (attr == "t1") curLabel.T1 = d;
                else if (attr == "title") curLabel.Title = s;
            }
            if (attr == "name" && s != null &&
                (elem == "wavetrack" || elem == "labeltrack" || elem == "notetrack" || elem == "timetrack"))
                TrackNames.Add(s);
        }

        // ---- writing -------------------------------------------------------

        // Names the document does not use yet are appended to the dictionary
        // with fresh ids. Nothing already in either blob moves, which is what
        // keeps a write surgical: the bytes Audacity wrote are the bytes that
        // stay.
        int EnsureName(string n)
        {
            int id;
            if (ids.TryGetValue(n, out id)) return id;
            if (maxId >= 65535) throw new Exception("the project name dictionary is full");
            id = ++maxId;
            ids[n] = id; names[id] = n;
            byte[] nb = Encoding.Unicode.GetBytes(n);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(Dict, 0, Dict.Length);
                ms.WriteByte(FT_Name);
                PutU16(ms, (ushort)id);
                PutU16(ms, (ushort)nb.Length);
                ms.Write(nb, 0, nb.Length);
                Dict = ms.ToArray();
            }
            return id;
        }

        static void PutU16(Stream s, ushort v) { s.WriteByte((byte)(v & 0xFF)); s.WriteByte((byte)(v >> 8)); }
        static void PutI32(Stream s, int v) { s.Write(BitConverter.GetBytes(v), 0, 4); }

        void PutStart(Stream s, string n) { s.WriteByte(FT_StartTag); PutU16(s, (ushort)EnsureName(n)); }
        void PutEnd(Stream s, string n) { s.WriteByte(FT_EndTag); PutU16(s, (ushort)EnsureName(n)); }
        void PutStr(Stream s, string n, string v)
        {
            byte[] b = Encoding.Unicode.GetBytes(v == null ? "" : v);
            s.WriteByte(FT_String); PutU16(s, (ushort)EnsureName(n)); PutI32(s, b.Length); s.Write(b, 0, b.Length);
        }
        void PutInt(Stream s, string n, int v) { s.WriteByte(FT_Int); PutU16(s, (ushort)EnsureName(n)); PutI32(s, v); }
        void PutBool(Stream s, string n, bool v) { s.WriteByte(FT_Bool); PutU16(s, (ushort)EnsureName(n)); s.WriteByte(v ? (byte)1 : (byte)0); }
        void PutDouble(Stream s, string n, double v)
        {
            s.WriteByte(FT_Double); PutU16(s, (ushort)EnsureName(n));
            s.Write(BitConverter.GetBytes(v), 0, 8);
            PutI32(s, 10);   // digits, as Audacity writes for t and t1
        }

        byte[] BuildLabelTrack(string name, IList<AuLabel> labels, int height)
        {
            using (MemoryStream s = new MemoryStream())
            {
                PutStart(s, "labeltrack");
                PutStr(s, "name", name);
                PutBool(s, "isSelected", false);
                PutInt(s, "height", height);
                PutBool(s, "minimized", false);
                PutInt(s, "numlabels", labels.Count);
                foreach (AuLabel l in labels)
                {
                    PutStart(s, "label");
                    PutDouble(s, "t", l.T);
                    PutDouble(s, "t1", l.T1);
                    PutStr(s, "title", l.Title);
                    PutEnd(s, "label");
                }
                PutEnd(s, "labeltrack");
                return s.ToArray();
            }
        }

        // Splice: drop the byte ranges of the label tracks being replaced, and
        // insert the new ones immediately before the closing project tag. Done
        // as one pass so no offset recorded during the parse is ever stale.
        public void ApplyLabelTracks(string[] trackNames, AuLabel[][] trackLabels, bool replaceExisting, int height)
        {
            if (trackNames.Length != trackLabels.Length)
                throw new Exception("track name and label arrays differ in length");

            List<byte[]> blocks = new List<byte[]>();
            for (int k = 0; k < trackNames.Length; k++)
                blocks.Add(BuildLabelTrack(trackNames[k], trackLabels[k], height));

            List<int[]> cuts = new List<int[]>();
            if (replaceExisting)
                foreach (AuLabelTrack lt in LabelTracks)
                    if (lt.DocStart >= 0 && lt.DocEnd > lt.DocStart) cuts.Add(new int[] { lt.DocStart, lt.DocEnd });
            cuts.Sort(delegate(int[] a, int[] b) { return a[0].CompareTo(b[0]); });

            int insertAt = ProjectEndOffset;
            using (MemoryStream outp = new MemoryStream())
            {
                int pos = 0;
                foreach (int[] cut in cuts)
                {
                    if (cut[0] > insertAt) throw new Exception("a label track sits after the closing project tag");
                    outp.Write(Doc, pos, cut[0] - pos);
                    pos = cut[1];
                }
                outp.Write(Doc, pos, insertAt - pos);
                foreach (byte[] b in blocks) outp.Write(b, 0, b.Length);
                outp.Write(Doc, insertAt, Doc.Length - insertAt);
                Doc = outp.ToArray();
            }

            // Re-read what was just built, so the caller's view of the document
            // is what is actually in it rather than what we meant to put there.
            AuProject fresh = Parse(Dict, Doc, Table, Path, RowId);
            LabelTracks = fresh.LabelTracks;
            TrackNames = fresh.TrackNames;
            ProjectEndOffset = fresh.ProjectEndOffset;
        }
    }

    // Just enough SQLite to move two BLOBs, over the copy Windows already ships.
    public static class AuSqlite
    {
        const int SQLITE_OK = 0, SQLITE_ROW = 100, SQLITE_DONE = 101;
        const int OPEN_READONLY = 0x00000001, OPEN_READWRITE = 0x00000002;
        static readonly IntPtr TRANSIENT = new IntPtr(-1);

        [DllImport("winsqlite3.dll", EntryPoint = "sqlite3_open_v2", CallingConvention = CallingConvention.Cdecl)]
        static extern int open_v2(byte[] file, out IntPtr db, int flags, IntPtr vfs);
        [DllImport("winsqlite3.dll", EntryPoint = "sqlite3_prepare_v2", CallingConvention = CallingConvention.Cdecl)]
        static extern int prepare_v2(IntPtr db, byte[] sql, int n, out IntPtr stmt, IntPtr tail);
        [DllImport("winsqlite3.dll", EntryPoint = "sqlite3_step", CallingConvention = CallingConvention.Cdecl)]
        static extern int step(IntPtr stmt);
        [DllImport("winsqlite3.dll", EntryPoint = "sqlite3_column_blob", CallingConvention = CallingConvention.Cdecl)]
        static extern IntPtr column_blob(IntPtr stmt, int col);
        [DllImport("winsqlite3.dll", EntryPoint = "sqlite3_column_bytes", CallingConvention = CallingConvention.Cdecl)]
        static extern int column_bytes(IntPtr stmt, int col);
        [DllImport("winsqlite3.dll", EntryPoint = "sqlite3_column_int64", CallingConvention = CallingConvention.Cdecl)]
        static extern long column_int64(IntPtr stmt, int col);
        [DllImport("winsqlite3.dll", EntryPoint = "sqlite3_bind_blob", CallingConvention = CallingConvention.Cdecl)]
        static extern int bind_blob(IntPtr stmt, int idx, byte[] val, int n, IntPtr destroy);
        [DllImport("winsqlite3.dll", EntryPoint = "sqlite3_finalize", CallingConvention = CallingConvention.Cdecl)]
        static extern int finalize_(IntPtr stmt);
        [DllImport("winsqlite3.dll", EntryPoint = "sqlite3_close", CallingConvention = CallingConvention.Cdecl)]
        static extern int close_(IntPtr db);
        [DllImport("winsqlite3.dll", EntryPoint = "sqlite3_errmsg", CallingConvention = CallingConvention.Cdecl)]
        static extern IntPtr errmsg(IntPtr db);
        [DllImport("winsqlite3.dll", EntryPoint = "sqlite3_changes", CallingConvention = CallingConvention.Cdecl)]
        static extern int changes(IntPtr db);

        static byte[] U8(string s) { return Encoding.UTF8.GetBytes(s + "\0"); }
        static string Err(IntPtr db)
        {
            IntPtr p = errmsg(db);
            if (p == IntPtr.Zero) return "unknown SQLite error";
            List<byte> b = new List<byte>();
            for (int i = 0; ; i++) { byte c = Marshal.ReadByte(p, i); if (c == 0) break; b.Add(c); }
            return Encoding.UTF8.GetString(b.ToArray());
        }
        static IntPtr Open(string path, bool write)
        {
            IntPtr db;
            int rc = open_v2(U8(path), out db, write ? OPEN_READWRITE : OPEN_READONLY, IntPtr.Zero);
            if (rc != SQLITE_OK)
            {
                string m = db == IntPtr.Zero ? ("SQLite error " + rc) : Err(db);
                if (db != IntPtr.Zero) close_(db);
                throw new Exception("Could not open " + path + ": " + m);
            }
            return db;
        }
        static IntPtr Prepare(IntPtr db, string sql)
        {
            IntPtr st;
            byte[] s = U8(sql);
            if (prepare_v2(db, s, s.Length, out st, IntPtr.Zero) != SQLITE_OK)
                throw new Exception("SQLite rejected a query on this project: " + Err(db));
            return st;
        }
        static byte[] Blob(IntPtr st, int col)
        {
            int n = column_bytes(st, col);
            IntPtr p = column_blob(st, col);
            if (p == IntPtr.Zero) return null;
            byte[] b = new byte[n];
            Marshal.Copy(p, b, 0, n);
            return b;
        }

        // -1 when the table is not there at all, which is how an .aup4unsaved
        // written by a newer build would announce itself.
        public static int CountRows(string path, string table)
        {
            IntPtr db = Open(path, false);
            IntPtr st = IntPtr.Zero;
            try
            {
                st = Prepare(db, "select count(*) from \"" + table + "\"");
                if (step(st) != SQLITE_ROW) return 0;
                return (int)column_int64(st, 0);
            }
            catch (Exception) { return -1; }
            finally { if (st != IntPtr.Zero) finalize_(st); close_(db); }
        }

        // { id, dict, doc } for the first row of the table, or null when empty.
        public static object[] ReadRow(string path, string table)
        {
            IntPtr db = Open(path, false);
            IntPtr st = IntPtr.Zero;
            try
            {
                st = Prepare(db, "select id, dict, doc from \"" + table + "\" order by id limit 1");
                if (step(st) != SQLITE_ROW) return null;
                return new object[] { column_int64(st, 0), Blob(st, 1), Blob(st, 2) };
            }
            finally { if (st != IntPtr.Zero) finalize_(st); close_(db); }
        }

        public static void WriteRow(string path, string table, long id, byte[] dict, byte[] doc)
        {
            IntPtr db = Open(path, true);
            IntPtr st = IntPtr.Zero;
            try
            {
                st = Prepare(db, "update \"" + table + "\" set dict = ?1, doc = ?2 where id = " + id);
                if (bind_blob(st, 1, dict, dict.Length, TRANSIENT) != SQLITE_OK) throw new Exception("binding the dictionary failed: " + Err(db));
                if (bind_blob(st, 2, doc, doc.Length, TRANSIENT) != SQLITE_OK) throw new Exception("binding the document failed: " + Err(db));
                if (step(st) != SQLITE_DONE) throw new Exception("writing the project document failed: " + Err(db));
                if (changes(db) != 1) throw new Exception("the project document row vanished while writing");
            }
            finally { if (st != IntPtr.Zero) finalize_(st); close_(db); }
        }
    }
}
'@
}

# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Opens an .aup3/.aup4 project file and hands back its document.
.DESCRIPTION
    Audacity keeps unsaved work in the autosave table and the last saved state
    in the project table, so autosave is preferred when it holds a row: that is
    what is on screen. Which one was used comes back as .Table.

    The file is always copied to a scratch directory and the copy is opened,
    never the original. Two reasons, both learned the hard way:

    - An .aup3/.aup4 is a WAL-mode database, so merely OPENING one - read-only,
      touching nothing - makes SQLite create "NAME.aup3-wal" and "NAME.aup3-shm"
      beside it. Read a folder of projects in place and you litter it with a
      side file per project, and they sync.
    - A project open in Audacity keeps its unsaved work in that -wal. Reading
      the database alone would silently return the last checkpointed state, so
      the -wal and -shm are copied along with it and the journal is honoured.

    Nothing here ever opens a file the user owns, which is why -NoCopy exists
    only for the writer, which has to.
#>
function Read-AudacityProjectFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Path,
        [switch]$NoCopy
    )

    if (-not (Test-Path -LiteralPath $Path)) { throw "There is no project file at:`n$Path" }

    $readFrom = $Path
    $scratch  = $null

    if (-not $NoCopy) {
        $scratch = Join-Path ([System.IO.Path]::GetTempPath()) ("AutoNyx-" + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $scratch -Force | Out-Null
        $leaf = [System.IO.Path]::GetFileName($Path)
        $readFrom = Join-Path $scratch $leaf
        Copy-Item -LiteralPath $Path -Destination $readFrom -Force
        foreach ($suffix in '-wal', '-shm') {
            $side = "$Path$suffix"
            if (Test-Path -LiteralPath $side) { Copy-Item -LiteralPath $side -Destination "$readFrom$suffix" -Force }
        }
    }

    try {
        $table = 'project'
        $autosaveRows = [AutoNyx.AuSqlite]::CountRows($readFrom, 'autosave')
        if ($autosaveRows -gt 0) { $table = 'autosave' }

        $row = [AutoNyx.AuSqlite]::ReadRow($readFrom, $table)
        if (-not $row) {
            throw "This project file holds no document. If it is open in Audacity, save it and try again."
        }

        $doc = [AutoNyx.AuProject]::Parse($row[1], $row[2], $table, $Path, [long]$row[0])
        return $doc
    }
    finally {
        if ($scratch -and (Test-Path -LiteralPath $scratch)) {
            Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

<#
.SYNOPSIS
    Every label in a project file, tagged with its label track's name.
.DESCRIPTION
    Shaped exactly like Get-AudacityAnnotations in lib\AudacityPipe.ps1 -
    Text, Start, End, Track, TrackNumber, sorted the same way - so the workbook
    writer cannot tell which of the two read the project.
#>
function Get-AudacityProjectFileAnnotations {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Path')] [string]$Path,
        [Parameter(Mandatory, ParameterSetName = 'Doc')]  $Document
    )

    if ($PSCmdlet.ParameterSetName -eq 'Path') { $Document = Read-AudacityProjectFile -Path $Path }

    $rows = New-Object System.Collections.Generic.List[object]
    $trackNumber = 0
    foreach ($track in $Document.LabelTracks) {
        $trackNumber++
        $name = if ($track.Name) { $track.Name } else { "Track $trackNumber" }
        foreach ($label in $track.Labels) {
            $rows.Add([pscustomobject]@{
                Text        = [string]$label.Title
                Start       = [double]$label.T
                End         = [double]$label.T1
                Track       = $name
                TrackNumber = $trackNumber
            })
        }
    }

    return , @($rows | Sort-Object Start, TrackNumber, End)
}

<#
.SYNOPSIS
    Writes label tracks into a project file that Audacity does not have open.
.DESCRIPTION
    Groups are what lib\SoundNouns.ps1 parses out of a caption file: Number,
    Name and Labels of Start/End/Text. Each becomes one label track called
    "2 Breathing", appended at the bottom in the order given - the same names
    and the same order the pipe route produces, so a project built either way
    reads back identically.

    -Replace deletes the label tracks already in the project first. Without it
    they are kept and the new ones land underneath.

    A copy of the file is put beside it before a single byte is written, and the
    write is verified by re-reading the file afterwards. Refuses outright if
    Audacity still has the project open, since Audacity would write its own
    picture of the project over ours on the next save.
#>
function Add-AudacityProjectFileLabelTracks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] $Groups,
        [switch]$Replace,
        [int]$TrackHeight = 73,
        [string]$BackupPath
    )

    if (-not (Test-Path -LiteralPath $Path)) { throw "There is no project file at:`n$Path" }

    # Writing underneath a running Audacity would be overwritten by its next
    # save at best, and tear the file at worst - so refuse while it is open.
    #
    # The test is whether anyone else holds the file open, NOT whether a -wal
    # sits beside it: a -wal outlives a crash, and any read-only open leaves an
    # empty one behind, so its mere presence proves nothing. Asking for the file
    # with no sharing is the question actually worth asking, because SQLite
    # opens with read/write sharing and so a live Audacity always fails this.
    try {
        $probe = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
        $probe.Close(); $probe.Dispose()
    } catch {
        throw ("Another program still has this project open:`n$Path`n`n" +
               "If that is Audacity, save and close the project there, then run this again.")
    }

    if (-not $BackupPath) {
        $stamp = Get-Date -Format 'yyyy-MM-dd HHmmss'
        $BackupPath = [System.IO.Path]::ChangeExtension($Path, $null).TrimEnd('.') +
                      " (before labels $stamp)" + [System.IO.Path]::GetExtension($Path)
    }
    Copy-Item -LiteralPath $Path -Destination $BackupPath -Force

    $document = Read-AudacityProjectFile -Path $Path
    $existingCount = $document.LabelTracks.Count

    $names  = New-Object System.Collections.Generic.List[string]
    $tracks = New-Object System.Collections.Generic.List[AutoNyx.AuLabel[]]
    foreach ($group in $Groups) {
        # "2 Breathing" where the caption file named the category, plain "2"
        # where it did not - the same rule ImportSoundNouns.ps1 follows.
        $trackName = if ($group.Name) { "$($group.Number) $($group.Name)" } else { [string]$group.Number }
        $names.Add($trackName)

        $labels = New-Object System.Collections.Generic.List[AutoNyx.AuLabel]
        foreach ($label in $group.Labels) {
            $entry = New-Object AutoNyx.AuLabel
            $entry.T     = [double]$label.Start
            $entry.T1    = [double]$label.End
            $entry.Title = [string]$label.Text
            $labels.Add($entry)
        }
        $tracks.Add($labels.ToArray())
    }

    $document.ApplyLabelTracks($names.ToArray(), $tracks.ToArray(), [bool]$Replace, $TrackHeight)
    [AutoNyx.AuSqlite]::WriteRow($Path, $document.Table, $document.RowId, $document.Dict, $document.Doc)

    # Read the file back rather than trusting the in-memory document: this is
    # the only check that what reached the disk is a project Audacity can parse.
    $after = Read-AudacityProjectFile -Path $Path
    $expected = if ($Replace) { $names.Count } else { $existingCount + $names.Count }
    if ($after.LabelTracks.Count -ne $expected) {
        throw ("The project now has $($after.LabelTracks.Count) label tracks, not the $expected expected. " +
               "The file as it was is at:`n$BackupPath")
    }

    return [pscustomobject]@{
        Path          = $Path
        BackupPath    = $BackupPath
        Table         = $document.Table
        TracksAdded   = $names.Count
        TracksRemoved = if ($Replace) { $existingCount } else { 0 }
        LabelTracks   = $after.LabelTracks.Count
        Labels        = ($after.LabelTracks | ForEach-Object { $_.Labels.Count } | Measure-Object -Sum).Sum
    }
}
