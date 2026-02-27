using System;

namespace MVP
{
    public interface ITaskpaneHostUI
    {
        event EventHandler RecognizeTankAssemlby;

        event EventHandler CreateTank;

        event EventHandler UpdateSettings;
    }
}
