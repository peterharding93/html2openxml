using HtmlToOpenXml.IO;
using NUnit.Framework;
using NUnit.Framework.Legacy;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests inline image (base 64).
    /// </summary>
    [TestFixture]
    public class DataUriTests
    {
        [Test]
        public void ParseInline()
        {
            // red dot
            string uri = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==";
            DataUri.TryCreate(uri, out DataUri result);
            ClassicAssert.IsNotNull(result);
        }

        [Test]
        public void ParseMultiline()
        {
            string uri = @"data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAoHBwgHBgoICAgLCgoLDhgQDg0NDh0VFhEYIx8l
JCIfIiEmKzcvJik0KSEiMEExNDk7Pj4+JS5ESUM8SDc9Pjv/wAALCABVAEABAREA/8QAHwAA
AQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQR
BRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RF
RkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ip
qrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/9oACAEB
AAA/APXNS1Gz0jT5r+/nWC2gXdJI3QD/AD2rz+P4wnVZ5YvDnhTU9VEfV1+UfU4DY/Guf1D4
561Y3LW0/heO0mXrHcSOGH4YFUW+Puufw6Pp4+pc/wBaVfj9rY+9o9gfozj+taOmfHDW9Tul
tbTwot5M3SO3lct+W01t3nxbv9DaI+IfBmoafHIcCQSBgfpkAE+2a7zRdasPEOlQ6lps4mt5
hwcYIPcEdiK4L473LReC7aBWI8+9UMPUBWP88V0fw0srey+H+kC3iVPOtxLIQOWduST61f8A
EvhPR/FdibXVbVZCB+7mXiSM+qt/TpXzv458Bah4KvgsrC4sZWIguV4z7MOxp/g74eal4qP2
yVhp+kx8zXs3C4HXbnqffoK+hvDHhzRvDmlR2+j2wjjdQWlYfvJfdj1NO8V6ba6v4W1KzvI1
eJ7dzyPusASGHuCAa85+AF4X0rWLIscRTxygem5SD/6CKf8AH9z/AGNo8Q/iuXP5KP8AGmaL
4c8UyQ6bpHiK9ls9LggCLHp1+kOzH/PU/eY9MBePcV0kvw10IwFoJtVvD2T+03AP1OapaP8A
BzRINSfUdVQXJLZjsw7NDGPcsdzn64HtUHiu6vIvif4b0i+VYPDrEGGNRiOWUA4DfRtmB05F
em1j+LroWfg/WLgnGyylI+uw4/WvKf2fnI1HWo+xhiP5Fv8AGr37QLf6Doi+ssp/RaofCDwz
4b8U6PqB1jTUvLy3uBmR5XB2MvHAPqGrtNS+FXheKwuJdNhuNMuFjZkmt7t12kDIJy2MVxGh
6n430SG3lbxTaSpLjyoNTWbypfQLMyBc+mGx71111e23xH0W98N6jaPpHiK0AmSCU8pIPuyI
38Snpkdj9DXS+Ctam13wvbXN2pS9i3W92h6rKh2tn64z+NYvxh1Iaf8ADy8jDAPeOkC/idx/
RTXE/s/Kf7T1luwhiH/jzf4VpftAITpWjSdhPIPzUf4VyvwR1oab40awkYCPUYTGM/31+Zf0
3D8a951qS0i0S9a/OLXyHEvPVSCCPqelc/4a0/xBqHhmHT/FMelz27wiOREDO7LjGG/h3e47
1w19Z3en2Wpi2mZ9W8E3KSWlw33pLNxuEbnuAufwGO9ei+FFjla81S1TbZ6v5N9GOwZ4wH/H
5QT9a8w+PetrNqen6HE2RbIZ5gP7zcKPwAJ/4FVv9n2LjXJsf88VH/j5rf8Ajjpst54IjuYo
y/2O6WR8D7qEFSfzIrwCyvJ9PvoL21cxz28iyRsOzA5FfTujeJbDxr4S+22drBfTBVM1hI4G
JBg7TnPcZBPB4rm9Q8V+J/CBEEXgiJbWZi8UFtdmQqzckBVXjnkgDHPWs2P+0NG8IeKfEniu
NbbUvESeTBZ/x/dKou3r/F06gLzXY2Op2fgL4eacdbnWGS2s0UxE/O77c7FHc54r5w13WLnX
9bu9Vuz+9upC5GeFHZR7AYH4V7b8CNMntPC15fTRlEvLn90SPvKoxn6ZJH4V1XjPxpo3hK3g
TWbaeeK+DoFjiVwQAMhgSPWuD8ReGvhvD4VtvFcmmaha2t8V8uO1kAfLZI+ViVHQ9KueHvhV
4evbG11zw94g1m0S4QPHJHKiuB3U4Ucg8Ee1XfEukz+E9FfU9R8e6+IEZUCpsZ3J7DOMnqev
QGub1PQPDD6LY+L9Z8Y684n+a081lM+QeijnBBHbgVjSweCpLiK51618ZiK4OIru92kP+mT+
Ga3Z9O+HGja0dM0zw5qPiPUIeZIoS0ix46huQOO/B9K77wd4107xJLcaZa6Zd6bPYIPMtp4Q
gjHQAY6fTiuL/aBA+waK3fzZR+i1neLYLjU/h94F8PWeGuL5VZFY4BIQAZ/77qX4I+KHsr+5
8J6gSm9mktlfgrIPvp+OM49j61T+LeqXnie4vGsSDo+gSLDLJniSdzg49cYx7c+tc5qmnanq
Pw80DV7eN57XTxNbzbRnyT5hYMR6EMBn2rudE+MWia7bQ6Z4u01IyWX9+F3wlgQQxXqvPPGa
47TtZ134V+L7uS4tFuUuchjJ925TOQ6OPXrnnryK9j8E+NvD3jCe5n06D7LqZjU3MUiASMq8
A7h94DOPbNcf+0BzYaKoBJ82U9PZapeHp/7R8b+A7PBZLHSRI2R0Yo5/otVviz4TvdI8X2ni
DRFkQ6jMoBh4ZLntj/e6/XNdX4i8JxeHvgnfaUvzzRxLPPJ1Ly71Zj+mPoBXE/DP4kWnhTTl
0nUoHazmuZHlmVSTFlVC8dwSGz3o+IEPhjxVqtjF4GtDc6lMx+0C0hZIyD0JBAAOe/510+i+
OfDEXhz/AIRbxtAkV1pYNrJHNC0qSbPlDKQDg8e3tVD4Q6BKPGmp67Y2lxbaJ5ckVo04IMgZ
1KjnqAByfXFezMiv95Q2PUZoCKOigY9BSkA9QDjmgqGBDAEHsaia0tnQxtbxMjdVKAg0W9pb
Wilba3ihB6iNAufyoltLaZw8tvFIw6FkBNSgADA4FLRRRRRRRX//2Q==
";
            DataUri.TryCreate(uri, out DataUri result);
            ClassicAssert.IsNotNull(result);
        }
    }
}
