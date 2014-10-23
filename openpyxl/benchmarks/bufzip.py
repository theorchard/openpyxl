# src http://stackoverflow.com/questions/297345/create-a-zip-file-from-a-generator-in-python

import os
import threading
from zipfile import *
import zlib, binascii, struct


class ZipEntryWriter(threading.Thread):
    def __init__(self, zf, zinfo, fileobj):
        self.zf = zf
        self.zinfo = zinfo
        self.fileobj = fileobj

        zinfo.file_size = 0
        zinfo.flag_bits = 0x00
        zinfo.header_offset = zf.fp.tell()

        zf._writecheck(zinfo)
        zf._didModify = True

        zinfo.CRC = 0
        zinfo.compress_size = compress_size = 0
        zf.fp.write(zinfo.FileHeader())

        super(ZipEntryWriter, self).__init__()

    def run(self):
        zinfo = self.zinfo
        zf = self.zf
        file_size = 0
        CRC = 0

        if zinfo.compress_type == ZIP_DEFLATED:
            cmpr = zlib.compressobj(zlib.Z_DEFAULT_COMPRESSION, zlib.DEFLATED, -15)
        else:
            cmpr = None
        while True:
            buf = self.fileobj.read(1024 * 8)
            if not buf:
                self.fileobj.close()
                break

            file_size = file_size + len(buf)
            #CRC = binascii.crc32(buf, CRC)
            CRC = binascii.crc32(buf, CRC) & 0xffffffff
            if cmpr:
                buf = cmpr.compress(buf)
                compress_size = compress_size + len(buf)

            zf.fp.write(buf)

        if cmpr:
            buf = cmpr.flush()
            compress_size = compress_size + len(buf)
            zf.fp.write(buf)
            zinfo.compress_size = compress_size
        else:
            zinfo.compress_size = file_size

        zinfo.CRC = CRC
        zinfo.file_size = file_size

        position = zf.fp.tell()
        zf.fp.seek(zinfo.header_offset + 14, 0)
        zf.fp.write(struct.pack("<LLL", zinfo.CRC, zinfo.compress_size, zinfo.file_size))
        zf.fp.seek(position, 0)
        zf.filelist.append(zinfo)
        zf.NameToInfo[zinfo.filename] = zinfo

class EnhZipFile(ZipFile, object):

    def _current_writer(self):
        return hasattr(self, 'cur_writer') and self.cur_writer or None

    def assert_no_current_writer(self):
        cur_writer = self._current_writer()
        if cur_writer and cur_writer.isAlive():
            raise ValueError('An entry is already started for name: %s' % cur_write.zinfo.filename)

    def write(self, filename, arcname=None, compress_type=None):
        self.assert_no_current_writer()
        super(EnhZipFile, self).write(filename, arcname, compress_type)

    def writestr(self, zinfo_or_arcname, bytes):
        self.assert_no_current_writer()
        super(EnhZipFile, self).writestr(zinfo_or_arcname, bytes)

    def close(self):
        self.finish_entry()
        super(EnhZipFile, self).close()

    def start_entry(self, zipinfo):
        """
        Start writing a new entry with the specified ZipInfo and return a
        file like object. Any data written to the file like object is
        read by a background thread and written directly to the zip file.
        Make sure to close the returned file object, before closing the
        zipfile, or the close() would end up hanging indefinitely.

        Only one entry can be open at any time. If multiple entries need to
        be written, make sure to call finish_entry() before calling any of
        these methods:
        - start_entry
        - write
        - writestr
        It is not necessary to explicitly call finish_entry() before closing
        zipfile.

        Example:
            zf = EnhZipFile('tmp.zip', 'w')
            w = zf.start_entry(ZipInfo('t.txt'))
            w.write("some text")
            w.close()
            zf.close()
        """
        self.assert_no_current_writer()
        r, w = os.pipe()
        self.cur_writer = ZipEntryWriter(self, zipinfo, os.fdopen(r, 'r'))
        self.cur_writer.start()
        return os.fdopen(w, 'w')

    def finish_entry(self, timeout=None):
        """
        Ensure that the ZipEntry that is currently being written is finished.
        Joins on any background thread to exit. It is safe to call this method
        multiple times.
        """
        cur_writer = self._current_writer()
        if not cur_writer or not cur_writer.isAlive():
            return
        cur_writer.join(timeout)


from openpyxl.compat import range
from openpyxl.xml.functions import XMLGenerator, start_tag, end_tag
from io import BytesIO
from zipfile import ZIP_DEFLATED
from functools import partial

def standard():
    archive = ZipFile("std.zip", "w", ZIP_DEFLATED)
    out = BytesIO()
    doc = XMLGenerator(out)
    for i in range(1000000):
        start_tag(doc, "test")
        end_tag(doc, "test")
    doc.endDocument()
    archive.writestr("sheet.xml", out.getvalue(), ZIP_DEFLATED)
    archive.close()


def buffered():
    archive = EnhZipFile("buffered.zip", "w", ZIP_DEFLATED)
    out = archive.start_entry(ZipInfo("sheet.xml"))
    #out.write = partial(out.write, compress_type=ZIP_DEFLATED)
    doc = XMLGenerator(out)
    for i in range(1000000):
        start_tag(doc, "test")
        end_tag(doc, "test")
    doc.endDocument()
    out.close()
    archive.close()


if __name__ == "__main__":
    zf = EnhZipFile('t.zip', 'w')
    import time
    from openpyxl.xml.functions import XMLGenerator, start_tag, end_tag
    w = zf.start_entry(ZipInfo('t.txt', time.localtime()[:6]))
    doc = XMLGenerator(w)
    start_tag(doc, "test")
    end_tag(doc, "test")
    doc.endDocument()
    w.write(u"Line1\n")
    w.write(u"Line2\n")
    w.close()
    zf.finish_entry()
    w = zf.start_entry(ZipInfo('p.txt', time.localtime()[:6]))
    w.write("Some text\n")
    w.close()
    zf.close()
