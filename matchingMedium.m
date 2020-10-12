PERM = 2.7;   %material permittivity
liquid_permittivity = 50;
desired_permittivity = 30;

%Do not touch below!
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

filename = 'infill_hexagon.xlsx';
infillTable = xlsread(filename);
infillTable = infillTable';

%infillTable's mm^2 is replaced with percentages
[ wallThicknessNumber sideLengthNumber] = size(infillTable);
percentages = infillTable( 2:wallThicknessNumber , 2:sideLengthNumber )*100/(47*47) ;
infillTable( 2:wallThicknessNumber , 2:sideLengthNumber ) = percentages;

extra_filename = 'extra_hexagon_infill.xlsx';
extra_infillTable = xlsread(extra_filename);
%extra_infillTable's mm^2 is replaced with percentages
[ wallThicknessNumber sideLengthNumber] = size(extra_infillTable);
percentages = extra_infillTable( 2:wallThicknessNumber , 2:sideLengthNumber )*100/(47*47) ;
extra_infillTable( 2:wallThicknessNumber , 2:sideLengthNumber ) = percentages;


catalogue{1,1} = xlsread('honeycombEpsilon_wallPermittivity1_wallThickness1mm.xlsx');
catalogue{1,2} = xlsread('honeycombEpsilon_wallPermittivity1_wallThickness1.5mm.xlsx');
catalogue{1,3} = xlsread('honeycombEpsilon_wallPermittivity1_wallThickness2mm.xlsx');
catalogue{1,4} = xlsread('honeycombEpsilon_wallPermittivity1_wallThickness3mm.xlsx');
catalogue{2,1} = xlsread('honeycombEpsilon_wallPermittivity2_wallThickness1mm.xlsx');
catalogue{2,2} = xlsread('honeycombEpsilon_wallPermittivity2_wallThickness1.5mm.xlsx');
catalogue{2,3} = xlsread('honeycombEpsilon_wallPermittivity2_wallThickness2mm.xlsx');
catalogue{2,4} = xlsread('honeycombEpsilon_wallPermittivity2_wallThickness3mm.xlsx');
catalogue{3,1} = xlsread('honeycombEpsilon_wallPermittivity3_wallThickness1mm.xlsx');
catalogue{3,2} = xlsread('honeycombEpsilon_wallPermittivity3_wallThickness1.5mm.xlsx');
catalogue{3,3} = xlsread('honeycombEpsilon_wallPermittivity3_wallThickness2mm.xlsx');
catalogue{3,4} = xlsread('honeycombEpsilon_wallPermittivity3_wallThickness3mm.xlsx');
catalogue{4,1} = xlsread('honeycombEpsilon_wallPermittivity4_wallThickness1mm.xlsx');
catalogue{4,2} = xlsread('honeycombEpsilon_wallPermittivity4_wallThickness1.5mm.xlsx');
catalogue{4,3} = xlsread('honeycombEpsilon_wallPermittivity4_wallThickness2mm.xlsx');
catalogue{4,4} = xlsread('honeycombEpsilon_wallPermittivity4_wallThickness3mm.xlsx');
catalogue{5,1} = xlsread('honeycombEpsilon_wallPermittivity5_wallThickness1mm.xlsx');
catalogue{5,2} = xlsread('honeycombEpsilon_wallPermittivity5_wallThickness1.5mm.xlsx');
catalogue{5,3} = xlsread('honeycombEpsilon_wallPermittivity5_wallThickness2mm.xlsx');
catalogue{5,4} = xlsread('honeycombEpsilon_wallPermittivity5_wallThickness3mm.xlsx');
catalogue{6,1} = xlsread('honeycombEpsilon_wallPermittivity6_wallThickness1mm.xlsx');
catalogue{6,2} = xlsread('honeycombEpsilon_wallPermittivity6_wallThickness1.5mm.xlsx');
catalogue{6,3} = xlsread('honeycombEpsilon_wallPermittivity6_wallThickness2mm.xlsx');
catalogue{6,4} = xlsread('honeycombEpsilon_wallPermittivity6_wallThickness3mm.xlsx');
catalogue{7,1} = xlsread('honeycombEpsilon_wallPermittivity7_wallThickness1mm.xlsx');
catalogue{7,2} = xlsread('honeycombEpsilon_wallPermittivity7_wallThickness1.5mm.xlsx');
catalogue{7,3} = xlsread('honeycombEpsilon_wallPermittivity7_wallThickness2mm.xlsx');
catalogue{7,4} = xlsread('honeycombEpsilon_wallPermittivity7_wallThickness3mm.xlsx');
catalogue{8,1} = xlsread('honeycombEpsilon_wallPermittivity8_wallThickness1mm.xlsx');
catalogue{8,2} = xlsread('honeycombEpsilon_wallPermittivity8_wallThickness1.5mm.xlsx');
catalogue{8,3} = xlsread('honeycombEpsilon_wallPermittivity8_wallThickness2mm.xlsx');
catalogue{8,4} = xlsread('honeycombEpsilon_wallPermittivity8_wallThickness3mm.xlsx');
catalogue{9,1} = xlsread('honeycombEpsilon_wallPermittivity9_wallThickness1mm.xlsx');
catalogue{9,2} = xlsread('honeycombEpsilon_wallPermittivity9_wallThickness1.5mm.xlsx');
catalogue{9,3} = xlsread('honeycombEpsilon_wallPermittivity9_wallThickness2mm.xlsx');
catalogue{9,4} = xlsread('honeycombEpsilon_wallPermittivity9_wallThickness3mm.xlsx');




flag = 0;
%catalogue generator
if( PERM > 8 )
    PERM = 8;
elseif(PERM < 1)
    PERM = 1;
end
if( round(PERM) == PERM )
    %do nothing
    flag = 1;
    newCatalogue = cell(4,1);
    for j = 1:4
        newCatalogue{j} = catalogue{PERM,j};
    end  
else
    higher_reference = ceil(PERM);
    lower_reference = floor(PERM); 
    increment = PERM - lower_reference;
    
    newCatalogue = cell(4,1);
    for j = 1:4
        higher_calalogue= catalogue{higher_reference,j};
        lower_catalogue= catalogue{lower_reference,j};

        addThis = ( increment / (higher_reference - lower_reference) ) * (higher_calalogue - lower_catalogue);
        newCatalogue{j} = lower_catalogue + addThis;
    end
end

%liquid permittivity range check
if( liquid_permittivity > 80 )
    liquid_permittivity = 80;
elseif( liquid_permittivity < 10 )
    liquid_permittivity = 10;
end

%liquid permittivity mid value insertion
if(  not(floor(liquid_permittivity/5) == (liquid_permittivity/5) ) )
    for j = 1:4
        thisCatalogue = newCatalogue{j};
        upper_liquidPermittivity = 5 * ceil(liquid_permittivity/5);
        lower_liquidPermittivity = 5 * floor(liquid_permittivity/5);

        [ rowNumber, columnNumber ] = size( thisCatalogue );

        for liquidCursor = 2:rowNumber
            catalogLiquidPermittivity = thisCatalogue( liquidCursor, 1 );
            if( isequal(catalogLiquidPermittivity , upper_liquidPermittivity))
                upper_liquidIndex = liquidCursor;
            end
            if( isequal(catalogLiquidPermittivity , lower_liquidPermittivity))
                lower_liquidIndex = liquidCursor;
            end
        end
        
        higher_permittivity_vector = thisCatalogue(upper_liquidIndex,:);
        lower_permittivity_vector = thisCatalogue(lower_liquidIndex,:);
        permittivity_difference = higher_permittivity_vector - lower_permittivity_vector;
        
        addThisToLower = ( (liquid_permittivity - lower_liquidPermittivity) /(upper_liquidPermittivity - lower_liquidPermittivity) ) * permittivity_difference;
        addThisToLower(1,1) =liquid_permittivity - lower_liquidPermittivity;
        
        intermediaryPermittivityList = lower_permittivity_vector + addThisToLower;
        thisCatalogue = [thisCatalogue ; zeros(1, columnNumber )];
        
        [ rowNumber, columnNumber ] = size( thisCatalogue );
        for i = rowNumber-1 : -1 : upper_liquidIndex 
            thisCatalogue(i+1,:) = thisCatalogue(i,:);
        end
        
        thisCatalogue(lower_liquidIndex +1 ,:) = intermediaryPermittivityList;
         
        newCatalogue{j} = thisCatalogue;
    end
end    


%detect liquid index
myLiquid_catalogueIndex = zeros(4,1);
for j = 1:4
    [ rowNumber, columnNumber ] = size( newCatalogue{j} );
    for liquid_index = 2 : rowNumber 
        searchLiquid = newCatalogue{j}( liquid_index, 1);
        if ( isequal(searchLiquid , liquid_permittivity) )
           myLiquid_catalogueIndex(j) = liquid_index ;
        end
    end
end


%linear estimation
a1 = zeros(4,1);
a0 = zeros(4,1);
for j= 1:4
    [ rowNumber, columnNumber ] = size( newCatalogue{j} );
    wallThickness = newCatalogue{j}(1,1 );
    
    [ infillrow, infillcolumn ] = size( infillTable );
    for infillTableIndex = 2:infillrow
        if isequal( infillTable(infillTableIndex,1) ,wallThickness )
            myInfillIndex = infillTableIndex;
        end
    end
    
    
    
    mySideLength = infillTable(1 , 2 : columnNumber )';
    myInfill = infillTable( myInfillIndex , 2 : columnNumber )';
    myCatalogue = newCatalogue{j}( myLiquid_catalogueIndex(j) , 2 : columnNumber )';
    
    
    [a0(j) a1(j)] = Linear_regression( myInfill, myCatalogue);
    x_model = myInfill(1,1)+5 : -0.5 : (myInfill(length(myInfill),1) - 5);
    y_model = a1(j) .*x_model + a0(j);
    hold on;
    plot( myInfill , myCatalogue,'x', x_model, y_model , 'MarkerSize',12 );
    xlabel({'Infill %'});
    ylabel({'Effective Dielectric Constant'});
    title(['Liquid Dielectric Constant '+ string(liquid_permittivity), 'Material Dielectric Constant '+ string(PERM) ])
end
hold off
legend('Wall Thickness: 1mm',' ','Wall Thickness: 1.5mm',' ','Wall Thickness: 2mm',' ','Wall Thickness: 3mm',' ');


[ rowNumberInfill, columnNumberInfill ] = size( infillTable );
fprintf( "Target Permittivity: %.2f  \n", desired_permittivity );
infill = zeros(1,4);
for j = 1 : 4
   infill(j) = ( desired_permittivity - a0(j) ) / a1(j);
      
   wall_thickness = newCatalogue{j}(1,1);
 
   wallThicknessIndexConfirm = infillTable(:,1) == wall_thickness;
   wallThicknessIndex = find( wallThicknessIndexConfirm , 1);
   
   lowSideLength = "Nan";
   highSideLength = "Check Infill Table";
   closerSidelengthValue = "Nan";
   for cursor = 2 : (columnNumberInfill-1)
        upperValue =  infillTable(wallThicknessIndex,cursor);
        lowerValue =  infillTable(wallThicknessIndex,cursor+1);

        if ( and( upperValue > infill(j) , lowerValue < infill(j) ) ) 
            lowSideLength = infillTable(1,cursor);
            highSideLength = infillTable(1,cursor+1);
            
            %infill is closer to...
            if( abs(lowerValue - infill(j) ) > abs(upperValue - infill(j) ) )
                closerSidelengthValue = lowSideLength;
            else
                closerSidelengthValue = highSideLength;   
            end
        end
   end

   fprintf( "for wall thickness %.2fmm, required infill percentage: %.2f (sidelength between %.1f - %.1fmm) closer to %.1fmm \n ", newCatalogue{j}(1,1) ,infill(j), lowSideLength, highSideLength, closerSidelengthValue );
end

%extra_infillTable
infill_average = sum( infill ) / 4 ; %average of infills
[ extra_rowInfill, extra_columnInfill ] = size( extra_infillTable );
fprintf("Further estimations by average infill %.2f:\n " , infill_average );
for j = 4:5
        
      lowSideLength = "Nan";
      highSideLength = "Check Infill Table";
      closerSidelengthValue = "Nan";
      for cursor = 2 : extra_columnInfill - 1
          extra_higherValue = extra_infillTable(j,cursor);
          extra_lowerValue = extra_infillTable(j,cursor+1);
          if and( extra_lowerValue < infill_average , extra_higherValue > infill_average )
              lowSideLength = extra_infillTable(1,cursor);
              highSideLength = extra_infillTable(1,cursor + 1);
              
              %infill is closer to...
              if( abs(extra_lowerValue - infill_average ) > abs(extra_higherValue - infill_average ) )
                  closerSidelengthValue = lowSideLength;
            else
                  closerSidelengthValue = highSideLength; 
             end
          end
 
      end
      fprintf("for wall thickness %.2fmm, with AVERAGE infill percentage  (sidelength between %.1f - %.1fmm) closer to %.1fmm \n ", j , lowSideLength, highSideLength, closerSidelengthValue );
end
