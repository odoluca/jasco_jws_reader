import io
import OleFileIO_PL as ofio
from struct import unpack


class JwsHeader:
    def __init__(self, channel_number, point_number,
                 x_for_first_point, x_for_last_point, x_increment,
                 data_size=None):
        self.channel_number = channel_number
        self.point_number = point_number
        self.x_for_first_point = x_for_first_point
        self.x_for_last_point = x_for_last_point
        self.x_increment = x_increment
        #only defined if this is the header of a v1.5 jws file
        self.data_size = data_size


def _unpack_ole_jws_header(data,format):
    # if len(data) < 96:
    #     raise _NotEnoughDataError, "DataInfo should be at least 96 bytes!"

    data_tuple = unpack(DATAINFO_FMT, data)
    print(data_tuple)
    # print(data_tuple[9:13])
    # I have only found spectra files with 1 channel
    # So we will only support 1 channel per file, at the moment
    return JwsHeader(data_tuple[3], data_tuple[5],
                     data_tuple[6], data_tuple[7], data_tuple[8])

file4="sample_CD_HT_Abs.jws"
file5="sample_CD_HT_Abs_2.jws"
DATAINFO_FMT = '<LLLLLLdddLLLLddddLLLLLLdddLLLLddd'
DATAINFO_FMT = '<LLLLLLdddLLLLLLdddLLLLLLdddLLdddd'
DATAINFO_FMT = '<LLLLLLdddLLLLLLddddLLddddLLdddd'
"""                 ^ ^^^^      ^^^^  ^^^^  ^^^^     """
# file1="sample_fluorescence.jws"
# file2="sample_fluorescence_2.jws"
# file3="sample_fluorescence_3.jws"
# DATAINFO_FMT = '<LLLLLLdddLLLLdddd'


def readheader(filename,format):
    f=open(filename,"rb")
    print(f.read(4))

    # print(file.read(100).decode('utf-8'))
    # f=io.StringIO(file)
    f.seek(0)
    oleobj=ofio.OleFileIO(f)
    str = oleobj.openstream('DataInfo')
    header_data = str.read()
    header_obj = _unpack_ole_jws_header(header_data,format)
    print(header_obj.point_number,header_obj.channel_number)


# readheader(file1,DATAINFO_FMT)
# readheader(file2,DATAINFO_FMT)
# readheader(file3,DATAINFO_FMT)
readheader(file4,DATAINFO_FMT)
readheader(file5,DATAINFO_FMT)
