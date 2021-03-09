function hyapp=hyconnect(Ruta_FileNameString,VisibleBoolean)
        hy=feval('actxserver','Hysys.application');
        
        if nargin<=1
              hy.Visible=1;
        else
              hy.Visible=VisibleBoolean;
        end
           try
              if nargin>0
                 hyy=invoke(hy.SimulationCases,'Open',Ruta_FileNameString);
                 invoke(hyy,'Activate');
               end
           catch
                disp('No se encuentra el archivo');
           end
             hyapp=hy;
