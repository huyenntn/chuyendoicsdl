using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AVDApplication
{
    public class Constants
    {
        public static class TableExport
        {
            public const string STT = "STT";
            public const string ID = "ID";
            public const string ID2 = "ID2";
            public const string GPNo = "Số GP";
            public const string MAU_GIAY_PHEP = "Mẫu giấy phép";
            public const string SO_THAM_CHIEU = "Số tham chiếu";
            public const string DO_LECH_F = "Độ lêch F CP";
            public const string TAN_SO = "Tần số";
            public const string BRAND_UU_TIEN = "Band ưu tiên";
            public const string DO_RONG_KENH = "Độ rộng kênh";
            public const string SO_KENH = "Số kênh";
            public const string TEN_KHACH_HANG = "Tên khách hàng";
            public const string HO_HIEU = "Hô hiệu";
            public const string KINH_DO = "Kinh độ";
            public const string VI_DO = "Vi do";
            public const string KINHDO_VIDO = "Kinh độ/Vĩ độ";
            public const string TEN_MAY = "Tên máy";

            // For Rohde and Schwarz
            public const string OFFSET_FREQ = "Offset kênh tần số";
            public const string DICH_VU = "Dịch Vụ";
            public const string KY_HIEU = "Ký hiệu";
            public const string DIEN_THOAI = "Điện thoại";
            public const string TEN_MA_DAT_NUOC = "Tên mã đất nước";
            public const string ZIP_CODE = "ZIPCODE";
            public const string TINH_THANH = "Tỉnh/Thành phố";
            public const string DUONG_PHO = "Đường/Phố";
            public const string HUONG_DAI_PHAT = "Hướng của đài phát";
            public const string KHOANG_CACH_DAI_PHAT = "Khoảng cách đến đài phát";
            public const string MIN_DO_LECH_FREQ = "Giá trị tối thiểu của độ lệch tần số";
            public const string BANG_THONG = "Băng thông";
            public const string MIN_DIEU_CHE = "Giá trị tối thiểu của điều chế";
            public const string DON_VI_DIEU_CHE = "Đơn vị điều chế";

            //For GEW

            public const string MUC_DICH_SU_DUNG = "Mục đích sử dụng";
            public const string DAI_LL = "Đài LL/Phương thức phát/Giờ LL";
            public const string DAI_LL2 = "Đài LL/Phương thức phát/Giờ LL2";

            public static class RSTABLE
            {
                public const string TRANSNAME = "'TRANSNAME";
                public const string FREQUENCY = "'FREQUENCY";
                public const string CHANNELOFS = "'CHANNELOFS";
                public const string SERVICE = "'SERVICE";
                public const string SIGNATURE = "'SIGNATURE";

                public const string CALLSIGN = "'CALLSIGN";
                public const string LICENSEE = "'LICENSEE";
                public const string TELEPHONE = "'TELEPHONE";
                public const string COUNTRY = "'COUNTRY";
                public const string ZIPCODE = "'ZIPCODE";

                public const string CITY = "'CITY";
                public const string STREET = "'STREET";
                public const string LONGITUDE = "'LONGITUDE";
                public const string LATITUDE = "'LATITUDE";
                public const string DIRECTION = "'DIRECTION";

                public const string DISTANCE = "'DISTANCE";
                public const string OFFSET = "'OFFSET";
                public const string BANDWIDTH = "'BANDWIDTH";
                public const string MODULATION = "'MODULATION";
                public const string MOD_UNIT = "'MOD_UNIT";
            }
            public static class GEWTABLE
            {
                public const string TRANSMITTER_EXTERNAL_ID = "Transmitter External ID";
                public const string FREQUENCY_EXTERNAL_ID = "Frequency External ID";
                public const string CENTRE_FREQUENCY = "Centre Frequency";
                public const string BANDWIDTH = "Bandwidth";
                public const string CHANNEL_SPACE = "Channel Space";
                public const string CHANNEL_NAME = "Channel Name";
                public const string NAME = "Name";
                public const string TYPE = "Type";
                public const string LATITUDE = "Latitude (deg)";
                public const string LONGITUDE = "Longitude (deg)";
                public const string COMMENT = "Comments";
            }

        }
        public static class ValueDAILL
        {
            public const string _16K0F3E = "16K0F3E";
            public const string _11K0F3E = "11K0F3E";
            public const string _6K50 = "6K50";
            public const string _6K5F3E = "6K5F3E";
        }
        public static class ValueConstant
        {
            public const string SPACE = " ";
            public const string RANDOM = "RANDOM";
            public const string NORMAL = "NORMAL";
            public const double LATITUDE = 20.0;
            public const double LONGTIDUDE = 105.0;

            public const string THTT = "PTTH tương tự";
            public const string THTS = "PTTH số";
            public const string DAI_TAU = "Ðài tàu";

            public const string HOURVALUE = "°";
        }

        public static class FreqAndStep
        {
            public class Frequency
            {
                public const string FREQ_HF_9_30 = "9K_30M";
                public const string FREQ_FM_47_50 = "47_50";
                public const string FREQ_FM_54_68 = "54_68";
                public const string FREQ_FM_87_108 = "87_108";
                public const string FREQ_HKHONG_108_138 = "108_138";
                public const string FREQ_VHF_138_174 = "138_174";
                public const string FREQ_VHF_174_230 = "174_230";
                public const string FREQ_UHF_400_463 = "400_470";
                public const string FREQ_UHF_470_806 = "470_806";
                public const string FREQ_CDMA_806_890 = "806_890";
                public const string FREQ_EGDSM_890_960 = "890_960";
                public const string FREQ_GSM_1800_1900 = "1800_1900";
                public const string FREQ_3G_2100_2170 = "2100_2170";
                public const string FREQ_3G_2620_2680 = "2620_2680";
            }
            public class FrequencyDisplay
            {
                public const string FREQ_HF_9_30 = "9KHz - 30MHz";
                public const string FREQ_FM_47_50 = "47MHz - 50MHz";
                public const string FREQ_FM_54_68 = "54MHz - 68MHz";
                public const string FREQ_FM_87_108 = "87MHz - 108MHz";
                public const string FREQ_HKHONG_108_138 = "108MHz - 138MHz";
                public const string FREQ_VHF_138_174 = "138MHz - 174MHz";
                public const string FREQ_VHF_174_230 = "174MHz - 230MHz";
                public const string FREQ_UHF_400_463 = "400MHz - 470MHz";
                public const string FREQ_UHF_470_806 = "470MHz - 806MHz";
                public const string FREQ_CDMA_806_890 = "806MHz - 890MHz";
                public const string FREQ_EGDSM_890_960 = "890MHz - 960MHz";
                public const string FREQ_GSM_1800_1900 = "1800MHz - 1900MHz";
                public const string FREQ_3G_2100_2170 = "2100MHz - 2170MHz";
                public const string FREQ_3G_2620_2680 = "2620MHz - 2680MHz";
            }

            public class FrequencyGEW
            {
                public const string FREQ_HF_9_30 = "9K_30M";
                public const string FREQ_TTKD_47_50 = "47_50";
                public const string FREQ_TTKD_54_68 = "54_68";
                public const string FREQ_PT_87_108 = "87_108";
                public const string FREQ_HK_108_137 = "108_137";
                public const string FREQ_DR_137_174 = "137_174";
                public const string FREQ_TH_174_230 = "174_230";
                public const string FREQ_DR_400_470 = "400_470";
                public const string FREQ_TH_470_790 = "470_790";
                public const string FREQ_TTDD_790_890 = "790_890";
                public const string FREQ_TTDD_890_960 = "890_960";
                public const string FREQ_TTDD_1710_1785 = "1710_1785";
                public const string FREQ_TTDD_1805_1880 = "1805_1880";
                public const string FREQ_TTDD_1920_1980 = "1920_1980";
                public const string FREQ_TTDD_2110_2170 = "2110_2170";
            }

            public class FrequencyGEWDisplay
            {
                public const string FREQ_HF_9_30 = "9KHz - 30MHz";
                public const string FREQ_TTKD_47_50 = "47MHz - 50MHz";
                public const string FREQ_TTKD_54_68 = "54MHz - 68MHz";
                public const string FREQ_PT_87_108 = "87MHz - 108MHz";
                public const string FREQ_HK_108_137 = "108MHz - 137MHz";
                public const string FREQ_DR_137_174 = "137MHz - 174MHz";
                public const string FREQ_TH_174_230 = "174MHz - 230MHz";
                public const string FREQ_DR_400_470 = "400MHz - 470MHz";
                public const string FREQ_TH_470_790 = "470MHz - 790MHz";
                public const string FREQ_TTDD_790_890 = "790MHz - 890MHz";
                public const string FREQ_TTDD_890_960 = "890MHz - 960MHz";
                public const string FREQ_TTDD_1710_1785 = "1710MHz - 1785MHz";
                public const string FREQ_TTDD_1805_1880 = "1805MHz - 1880MHz";
                public const string FREQ_TTDD_1920_1980 = "1920MHz - 1980MHz";
                public const string FREQ_TTDD_2110_2170 = "2110MHz - 2170MHz";
            }

            public class Step
            {
                public const string STEP_1 = "1";
                public const string STEP_3 = "3";
                public const string STEP_5 = "5";
                public const string STEP_6_25 = "6.25";
                public const string STEP_10 = "10";
                public const string STEP_12_5 = "12.5";
                public const string STEP_15 = "15";

                public const string STEP_20 = "20";
                public const string STEP_25 = "25";
                public const string STEP_30 = "30";
                public const string STEP_50 = "50";
                public const string STEP_100 = "100";
                public const string STEP_plus100 = "+100";
                public const string STEP_minus100 = "-100";
                public const string STEP_200 = "200";

            }
        }
    }
}
