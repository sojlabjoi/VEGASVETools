using System;
using System.Windows.Forms;
using System.Collections.Generic;

using Sony.Vegas;

public class EntryPoint
{
    Vegas vegas;

    int tracks;

    Dictionary<string, int> videoTracks = new Dictionary<string, int>();

    public enum ResampleMode
    {
        Smart,
        Force,
        Disable
    }

    public enum FlipStyle
    {
        None,
        Horizontal,
        Vertical,
        ZigZag,
        Arpeggio1,
        Arpeggio2,
    }

    public struct Options
    {
        public int selected;

        public FlipStyle flip;
        public ResampleMode resample;
    }

    public void FromVegas(Vegas veg)
    {
        vegas = veg;

        IndexVideoTracks();

        if (videoTracks.Count == 0)
        {
            if (veg.Project.FilePath == null)
            {
                MessageBox.Show(
                    "Cannot find video track(s) in Untitled.",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show(String.Format(
                    "Cannot find video track(s) in {0}.",
                    System.IO.Path.GetFileName(veg.Project.FilePath)),
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return;
        }

        Options? opt = ScriptForm.GetOptions(videoTracks);
        if (opt != null && opt.HasValue)
        {
            ApplyChanges(opt.Value);
        }
        return;
    }

    public void IndexVideoTracks()
    {
        tracks = vegas.Project.Tracks.Count;
        for (int i = 0; i < tracks; i++)
        {
            Track t = vegas.Project.Tracks[i];
            if (t.IsVideo())
            {
                if (t.Name != null)
                {
                    videoTracks.Add(String.Format("Track {0} [{1}] ({2} events)",
                    i + 1, t.Name, t.Events.Count), i);
                }
                else
                {
                    videoTracks.Add(String.Format("Track {0} ({1} events)",
                    i + 1, t.Events.Count), i);
                }
            }
        }
    }

    public void ApplyChanges(Options opt)
    {
        int count = 0;
        int flipped = 0;
        int resampled = 0;

        Track t = vegas.Project.Tracks[opt.selected];
        if (t.Events.Count == 0)
        {
            MessageBox.Show(String.Format(
                "Cannot find video event(s) in track {0}.",
                opt.selected + 1),
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        foreach (TrackEvent te in t.Events)
        {
            VideoEvent ve = (VideoEvent)te;
            switch (opt.resample)
            {
                case ResampleMode.Smart:
                    if (ve.ResampleMode != VideoResampleMode.Smart)
                    {
                        ve.ResampleMode = VideoResampleMode.Smart;
                        resampled++;
                    }
                    break;
                case ResampleMode.Force:
                    if (ve.ResampleMode != VideoResampleMode.Force)
                    {
                        ve.ResampleMode = VideoResampleMode.Force;
                        resampled++;
                    }
                    break;
                case ResampleMode.Disable:
                    if (ve.ResampleMode != VideoResampleMode.Disable)
                    {
                        ve.ResampleMode = VideoResampleMode.Disable;
                        resampled++;
                    }
                    break;
            }

            VideoMotionKeyframe vmk = ve.VideoMotion.Keyframes[0];

            VideoMotionVertex tl4 = vmk.TopLeft;
            VideoMotionVertex tr3 = vmk.TopRight;
            VideoMotionVertex br2 = vmk.BottomRight;
            VideoMotionVertex bl1 = vmk.BottomLeft;

            VideoMotionBounds vmb4 = new VideoMotionBounds(
                bl1, br2, tr3, tl4);
            VideoMotionBounds vmb3 = new VideoMotionBounds(
                br2, bl1, tl4, tr3);
            VideoMotionBounds vmb2 = new VideoMotionBounds(
                tr3, tl4, bl1, br2);
            if (count % 2 == 1)
            {
                switch (opt.flip)
                {
                    case FlipStyle.None:
                        break;
                    case FlipStyle.Horizontal:
                        ve.VideoMotion.Keyframes[0].Bounds = vmb2;
                        flipped++;
                        break;
                    case FlipStyle.Vertical:
                        ve.VideoMotion.Keyframes[0].Bounds = vmb4;
                        flipped++;
                        break;
                }
            }
            if (count % 4 == 1)
            {
                switch (opt.flip)
                {
                    case FlipStyle.ZigZag:
                        ve.VideoMotion.Keyframes[0].Bounds = vmb2;
                        flipped++;
                        break;
                    case FlipStyle.Arpeggio1:
                        ve.VideoMotion.Keyframes[0].Bounds = vmb2;
                        flipped++;
                        break;
                }
            }
            if (count % 4 == 2)
            {
                switch (opt.flip)
                {
                    case FlipStyle.ZigZag:
                        ve.VideoMotion.Keyframes[0].Bounds = vmb4;
                        flipped++;
                        break;
                    case FlipStyle.Arpeggio1:
                        ve.VideoMotion.Keyframes[0].Bounds = vmb3;
                        flipped++;
                        break;
                }
            }
            if (count % 4 == 3)
            {
                switch (opt.flip)
                {
                    case FlipStyle.ZigZag:
                        ve.VideoMotion.Keyframes[0].Bounds = vmb3;
                        flipped++;
                        break;
                    case FlipStyle.Arpeggio1:
                        ve.VideoMotion.Keyframes[0].Bounds = vmb4;
                        flipped++;
                        break;
                }
            }
            count++;
        }

        MessageBox.Show(String.Format(
        "Successfully applied changes to {0} events ({1} flipped | {2} resampled)",
        count, flipped, resampled),
        "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        return;
    }

    class ScriptForm
    {
        public static Options? GetOptions(Dictionary<string, int> videoTracks)
        {
            Form f = new Form();
            f.Text = "VEGAS VideoEvent Tools";
            f.Size = new System.Drawing.Size(420, 195);
            f.FormBorderStyle = FormBorderStyle.FixedSingle;
            f.MaximizeBox = false;
            f.MinimizeBox = false;

            ComboBox cbList = new ComboBox();
            cbList.Anchor = AnchorStyles.Top;
            cbList.Location = new System.Drawing.Point(0, 5);
            cbList.Width = f.Width - 16;
            cbList.DropDownStyle = ComboBoxStyle.DropDownList;
            foreach (KeyValuePair<string, int> t in videoTracks)
            {
                cbList.Items.Add(t.Key);
                cbList.SelectedIndex = 0;
            }

            GroupBox gbProp = new GroupBox();
            InitialiseProperties(f, cbList, gbProp);

            GroupBox gbPanCrop = new GroupBox();
            InitialisePanCrop(f, cbList, gbProp, gbPanCrop);

            Button bApply = new Button();
            bApply.Text = "Apply";
            bApply.Anchor = AnchorStyles.Bottom;
            bApply.Location = new System.Drawing.Point(
                (f.Width / 2) - (bApply.Width / 2), f.Height - 65);
            bApply.DialogResult = DialogResult.OK;

            f.Controls.Add(cbList);
            f.Controls.Add(gbProp);
            f.Controls.Add(gbPanCrop);
            f.Controls.Add(bApply);

            f.ShowDialog();
            if (f.DialogResult == DialogResult.OK)
            {
                Options opt = new Options();
                opt.selected = videoTracks[cbList.Text];
                foreach (Control ctrlProp in gbProp.Controls)
                {
                    RadioButton rbResample = (RadioButton)ctrlProp;
                    if (rbResample.Checked)
                    {
                        opt.resample = (ResampleMode)rbResample.Tag;
                    }
                }
                foreach (Control ctrlPanCrop in gbPanCrop.Controls)
                {
                    RadioButton rbFlip = (RadioButton)ctrlPanCrop;
                    if (rbFlip.Checked)
                    {
                        opt.flip = (FlipStyle)rbFlip.Tag;
                    }
                }
                return opt;
            }
            return null;
        }

        static void InitialiseProperties(Form f, ComboBox cbList, GroupBox gbProp)
        {
            gbProp.Text = "Event Properties";
            gbProp.Anchor = AnchorStyles.Top;
            gbProp.Top = cbList.Bottom + 10;
            gbProp.Width = f.Width / 3;
            gbProp.Height = 100;

            RadioButton rbSmart = new RadioButton();
            rbSmart.Text = "Smart Resample";
            rbSmart.Location = new System.Drawing.Point(
                gbProp.Left + 10, gbProp.Top - 20);
            rbSmart.Width = gbProp.Width - 20;
            rbSmart.BackColor = System.Drawing.Color.Transparent;
            rbSmart.Tag = ResampleMode.Smart;
            rbSmart.Checked = true;

            RadioButton rbForce = new RadioButton();
            rbForce.Text = "Forced Resample";
            rbForce.Location = new System.Drawing.Point(
                rbSmart.Left, rbSmart.Bottom);
            rbForce.Width = gbProp.Width - 20;
            rbForce.BackColor = System.Drawing.Color.Transparent;
            rbForce.Tag = ResampleMode.Force;

            RadioButton rbDisable = new RadioButton();
            rbDisable.Text = "Disable Resample";
            rbDisable.Location = new System.Drawing.Point(
                rbForce.Left, rbForce.Bottom);
            rbDisable.Width = gbProp.Width - 20;
            rbDisable.BackColor = System.Drawing.Color.Transparent;
            rbDisable.Tag = ResampleMode.Disable;

            gbProp.Controls.Add(rbSmart);
            gbProp.Controls.Add(rbForce);
            gbProp.Controls.Add(rbDisable);
        }

        static void InitialisePanCrop(Form f, ComboBox cbList, GroupBox gbProp, GroupBox gbPanCrop)
        {
            gbPanCrop.Text = "Event Pan/Crop";
            gbPanCrop.Anchor = AnchorStyles.Top;
            gbPanCrop.Top = cbList.Bottom + 10;
            gbPanCrop.Left = f.Width / 3 + 33;
            gbPanCrop.Width = f.Width / 2 + 22;
            gbPanCrop.Height = 100;

            RadioButton rbNone = new RadioButton();
            rbNone.Text = "None";
            rbNone.Location = new System.Drawing.Point(
                gbPanCrop.Left / 16 + 1, gbPanCrop.Top - 20);
            rbNone.Width = gbPanCrop.Width / 2 - 12;
            rbNone.BackColor = System.Drawing.Color.Transparent;
            rbNone.Tag = FlipStyle.None;
            rbNone.Checked = true;

            RadioButton rbHoriz = new RadioButton();
            rbHoriz.Text = "Flip Horizontally";
            rbHoriz.Location = new System.Drawing.Point(
                rbNone.Left, rbNone.Bottom);
            rbHoriz.Width = gbPanCrop.Width / 2 - 12;
            rbHoriz.BackColor = System.Drawing.Color.Transparent;
            rbHoriz.Tag = FlipStyle.Horizontal;

            RadioButton rbVert = new RadioButton();
            rbVert.Text = "Flip Vertically";
            rbVert.Location = new System.Drawing.Point(
                rbHoriz.Left, rbHoriz.Bottom);
            rbVert.Width = gbPanCrop.Width / 2 - 12;
            rbVert.BackColor = System.Drawing.Color.Transparent;
            rbVert.Tag = FlipStyle.Vertical;

            RadioButton rbZigZag = new RadioButton();
            rbZigZag.Text = "Zig-Zag";
            rbZigZag.Location = new System.Drawing.Point(
                gbPanCrop.Left / 16 + 116, gbPanCrop.Top - 20);
            rbZigZag.Width = gbPanCrop.Width / 2;
            rbZigZag.BackColor = System.Drawing.Color.Transparent;
            rbZigZag.Tag = FlipStyle.ZigZag;

            RadioButton rbArp1 = new RadioButton();
            rbArp1.Text = "Arpeggio 1";
            rbArp1.Location = new System.Drawing.Point(
                rbZigZag.Left, rbZigZag.Bottom);
            rbArp1.Width = gbPanCrop.Width / 2;
            rbArp1.BackColor = System.Drawing.Color.Transparent;
            rbArp1.Tag = FlipStyle.Arpeggio1;

            RadioButton rbArp2 = new RadioButton();
            rbArp2.Text = "Arpeggio 2";
            rbArp2.Location = new System.Drawing.Point(
                rbArp1.Left, rbArp1.Bottom);
            rbArp2.Width = gbPanCrop.Width / 2;
            rbArp2.BackColor = System.Drawing.Color.Transparent;
            rbArp2.Tag = FlipStyle.Arpeggio2;

            gbPanCrop.Controls.Add(rbNone);
            gbPanCrop.Controls.Add(rbHoriz);
            gbPanCrop.Controls.Add(rbVert);
            gbPanCrop.Controls.Add(rbZigZag);
            gbPanCrop.Controls.Add(rbArp1);
            //gbPanCrop.Controls.Add(rbArp2);
        }
    }
}
