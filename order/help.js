SetupResize()
  {
  window.onResize = ResizePage;
  }

ResizePage()
  {
  window.;
  }

AutoSizeImg(ii)
  {
  if (ii.width > 300 && document.body.clientWidth < ii.width) 
    {
    if (document.body.clientWidth > 400) 
      {
      ii.width=document.body.clientWidth - 15;
      }
     else
      {
      ii.width = 385;
      }
    }
  }