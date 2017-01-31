﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Input;
using System.Windows.Controls.Primitives;

namespace PowerVBA.UserControls
{
    [TemplatePart(Name = "SimpleButton", Type = typeof(Border))]
    [TemplatePart(Name = "ExButton", Type = typeof(Border))]
    public class ImageButton : Control
    {
        public ImageButton()
        {
            this.Style = FindResource("ImageButtonStyle") as Style;
        }


        public static DependencyProperty ContentProperty = DependencyProperty.Register(nameof(Content), typeof(string), typeof(ImageButton));
        public static DependencyProperty BackImageProperty = DependencyProperty.Register(nameof(BackImage), typeof(ImageSource), typeof(ImageButton));
        public static DependencyProperty ButtonModeProperty = DependencyProperty.Register(nameof(buttonMode), typeof(ButtonMode), typeof(ImageButton));

        public ImageSource BackImage
        {
            get { return (ImageSource)GetValue(BackImageProperty); }
            set { SetValue(BackImageProperty, value); }
        }
        public string Content
        {
            get { return (string)GetValue(ContentProperty); }
            set { SetValue(ContentProperty, value); }
        }

        public ButtonMode buttonMode
        {
            get { return (ButtonMode)GetValue(ButtonModeProperty); }
            set { SetValue(ButtonModeProperty, value); }
        }

        private Border SimpleBtn;
        private Border ExBtn;


        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();

            SimpleBtn = GetTemplateChild("SimpleButton") as Border;
            ExBtn = GetTemplateChild("ExButton") as Border;

            SimpleBtn.MouseLeftButtonDown += SimpleBtn_LeftMouseDown;
            SimpleBtn.MouseUp += SimpleBtn_LeftButtonUp;
            if (ExBtn != null)
            {
                ExBtn.MouseLeftButtonDown += ExBtn_LeftButtonDown;
                ExBtn.MouseLeftButtonUp += ExBtn_LeftButtonUp;
            }
            
            
        }

        bool ExDown, SimpleDown;

        private void ExBtn_LeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            
            ExDown = true;
        }

        private void ExBtn_LeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (ExDown && ExButtonClicked != null) ExButtonClicked();
            ExDown = false;
        }



        private void SimpleBtn_LeftMouseDown(object sender, MouseButtonEventArgs e)
        {
            SimpleDown = true;
        }
        private void SimpleBtn_LeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (SimpleDown && SimpleButtonClicked != null) SimpleButtonClicked();
            SimpleDown = false;
        }
        public delegate void BlankEventHandler();
        public event BlankEventHandler SimpleButtonClicked;
        public event BlankEventHandler ExButtonClicked;



        public enum ButtonMode
        {
            /// <summary>
            /// 기본 버튼입니다.
            /// </summary>
            Default,
            /// <summary>
            /// 버튼과 자세히 보기 두개의 버튼이 있습니다.
            /// </summary>
            ButtonWithDetails,
            /// <summary>
            /// 오직 자세히 보기 버튼만 있습니다.
            /// </summary>
            OnlyDetails
        }
    }
}
