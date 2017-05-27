Imports Microsoft.Kinect

Public Class Form1

    Dim kinz As KinectSensor
    Dim imagez As ColorImageFrame
    Dim skeltonz As SkeletonFrame
    Dim piccolor As Bitmap = New Bitmap(640, 480, Imaging.PixelFormat.Format32bppRgb)
    Dim gfx As Graphics = Graphics.FromImage(piccolor)

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        For Each potentialSensor In KinectSensor.KinectSensors
            If potentialSensor.Status = KinectStatus.Connected Then
                kinz = potentialSensor
                Exit For
            End If
        Next

        kinz.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30)
        kinz.SkeletonStream.Enable()
        AddHandler kinz.ColorFrameReady, AddressOf colorready
        AddHandler kinz.SkeletonFrameReady, AddressOf skeltonready
        kinz.Start()

        kinz.SkeletonStream.TrackingMode = SkeletonTrackingMode.Seated
        kinz.ElevationAngle = -10
    End Sub

    Private Sub colorready(ByVal sender As Object, ByVal e As ColorImageFrameReadyEventArgs)
        imagez = e.OpenColorImageFrame
    End Sub

    Private Sub skeltonready(ByVal sender As Object, ByVal e As SkeletonFrameReadyEventArgs)
        skeltonz = e.OpenSkeletonFrame
    End Sub

    Public Sub colormethod()

        Dim pixz(kinz.ColorStream.FramePixelDataLength - 1) As Byte
        If imagez IsNot Nothing Then
            imagez.CopyPixelDataTo(pixz)

            Dim rect As New Rectangle(0, 0, piccolor.Width, piccolor.Height)
            Dim bmpData As System.Drawing.Imaging.BitmapData = piccolor.LockBits(rect, _
                Drawing.Imaging.ImageLockMode.ReadWrite, Imaging.PixelFormat.Format32bppRgb)
            Dim ptr As IntPtr = bmpData.Scan0
            Dim bytes As Integer = bmpData.Stride * piccolor.Height
            Dim rgbValues(bytes - 1) As Byte
            System.Runtime.InteropServices.Marshal.Copy(ptr, rgbValues, 0, bytes)

            Dim secondcounter As Integer
            Dim tempred As Integer
            Dim tempblue As Integer
            Dim tempgreen As Integer
            Dim tempalpha As Integer
            Dim tempx As Integer
            Dim tempy As Integer
            secondcounter = 0

            While secondcounter < rgbValues.Length
                tempblue = rgbValues(secondcounter)
                tempgreen = rgbValues(secondcounter + 1)
                tempred = rgbValues(secondcounter + 2)
                tempalpha = rgbValues(secondcounter + 3)
                tempalpha = 225

                tempy = ((secondcounter * 0.25) / 640)
                tempx = ((secondcounter * 0.25) - (tempy * 640))
                If tempx < 0 Then
                    tempx = tempx + 640
                End If

                tempred = pixz(secondcounter + 2)
                tempgreen = pixz(secondcounter + 1)
                tempblue = pixz(secondcounter + 0)

                rgbValues(secondcounter) = tempblue
                rgbValues(secondcounter + 1) = tempgreen
                rgbValues(secondcounter + 2) = tempred
                rgbValues(secondcounter + 3) = tempalpha

                secondcounter = secondcounter + 4
            End While

            System.Runtime.InteropServices.Marshal.Copy(rgbValues, 0, ptr, bytes)

            piccolor.UnlockBits(bmpData)
        End If
    End Sub

    Public Sub skeltonmethod()

        Dim skeltons(-1) As Skeleton
        If skeltonz IsNot Nothing Then
            skeltons = New Skeleton(skeltonz.SkeletonArrayLength - 1) {}
            skeltonz.CopySkeletonDataTo(skeltons)
        End If

        Dim penz As Pen = New Pen(Brushes.LimeGreen, 3)

        If skeltons.Length <> 0 Then

            For Each skel As Skeleton In skeltons

                'Right Arm
                Dim shoulderright As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.ShoulderRight).Position, DepthImageFormat.Resolution640x480Fps30)
                Dim elbowright As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.ElbowRight).Position, DepthImageFormat.Resolution640x480Fps30)
                Dim wristright As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.WristRight).Position, DepthImageFormat.Resolution640x480Fps30)
                Dim handright As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.HandRight).Position, DepthImageFormat.Resolution640x480Fps30)
                gfx.DrawLine(penz, New Point(shoulderright.X, shoulderright.Y), New Point(elbowright.X, elbowright.Y))
                gfx.DrawLine(penz, New Point(elbowright.X, elbowright.Y), New Point(wristright.X, wristright.Y))
                gfx.DrawLine(penz, New Point(wristright.X, wristright.Y), New Point(handright.X, handright.Y))

                'Left Arm
                Dim shoulderleft As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.ShoulderLeft).Position, DepthImageFormat.Resolution640x480Fps30)
                Dim elbowleft As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.ElbowLeft).Position, DepthImageFormat.Resolution640x480Fps30)
                Dim wristleft As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.WristLeft).Position, DepthImageFormat.Resolution640x480Fps30)
                Dim handleft As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.HandLeft).Position, DepthImageFormat.Resolution640x480Fps30)
                gfx.DrawLine(penz, New Point(shoulderleft.X, shoulderleft.Y), New Point(elbowleft.X, elbowleft.Y))
                gfx.DrawLine(penz, New Point(elbowleft.X, elbowleft.Y), New Point(wristleft.X, wristleft.Y))
                gfx.DrawLine(penz, New Point(wristleft.X, wristleft.Y), New Point(handleft.X, handleft.Y))

                'Right Leg
                Dim FootRight As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.FootRight).Position, DepthImageFormat.Resolution640x480Fps30)
                Dim AnkleRight As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.AnkleRight).Position, DepthImageFormat.Resolution640x480Fps30)
                Dim KneeRight As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.KneeRight).Position, DepthImageFormat.Resolution640x480Fps30)
                Dim HipRight As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.HipRight).Position, DepthImageFormat.Resolution640x480Fps30)
                gfx.DrawLine(penz, New Point(HipRight.X, HipRight.Y), New Point(KneeRight.X, KneeRight.Y))
                gfx.DrawLine(penz, New Point(KneeRight.X, KneeRight.Y), New Point(AnkleRight.X, AnkleRight.Y))
                gfx.DrawLine(penz, New Point(AnkleRight.X, AnkleRight.Y), New Point(FootRight.X, FootRight.Y))

                'Left Leg
                Dim FootLeft As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.FootLeft).Position, DepthImageFormat.Resolution640x480Fps30)
                Dim AnkleLeft As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.AnkleLeft).Position, DepthImageFormat.Resolution640x480Fps30)
                Dim KneeLeft As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.KneeLeft).Position, DepthImageFormat.Resolution640x480Fps30)
                Dim HipLeft As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.HipLeft).Position, DepthImageFormat.Resolution640x480Fps30)
                gfx.DrawLine(penz, New Point(HipLeft.X, HipLeft.Y), New Point(KneeLeft.X, KneeLeft.Y))
                gfx.DrawLine(penz, New Point(KneeLeft.X, KneeLeft.Y), New Point(AnkleLeft.X, AnkleLeft.Y))
                gfx.DrawLine(penz, New Point(AnkleLeft.X, AnkleLeft.Y), New Point(FootLeft.X, FootLeft.Y))

                'Body
                Dim head As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.Head).Position, DepthImageFormat.Resolution640x480Fps30)
                Dim shouldercenter As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.ShoulderCenter).Position, DepthImageFormat.Resolution640x480Fps30)
                Dim Spine As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.Spine).Position, DepthImageFormat.Resolution640x480Fps30)
                Dim HipCenter As DepthImagePoint = kinz.MapSkeletonPointToDepth(skel.Joints(JointType.HipCenter).Position, DepthImageFormat.Resolution640x480Fps30)
                gfx.DrawLine(penz, New Point(head.X, head.Y), New Point(shouldercenter.X, shouldercenter.Y))
                gfx.DrawLine(penz, New Point(shouldercenter.X, shouldercenter.Y), New Point(shoulderleft.X, shoulderleft.Y))
                gfx.DrawLine(penz, New Point(shouldercenter.X, shouldercenter.Y), New Point(Spine.X, Spine.Y))
                gfx.DrawLine(penz, New Point(HipCenter.X, HipCenter.Y), New Point(Spine.X, Spine.Y))
                gfx.DrawLine(penz, New Point(HipCenter.X, HipCenter.Y), New Point(HipRight.X, HipRight.Y))
                gfx.DrawLine(penz, New Point(HipCenter.X, HipCenter.Y), New Point(HipLeft.X, HipLeft.Y))
            Next
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        colormethod()
        skeltonmethod()
        PictureBox1.Image = piccolor
    End Sub
End Class
