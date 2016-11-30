<?php
/*
 * Copyright (C) 1996-2016 YONGF Inc.All Rights Reserved.
 * Scott Wang blog.54yongf.com | blog.csdn.net/yongf2014	
 * 文件名：main.php @ CitePro						
 * 描述：
 * 
 * 修改历史
 * 版本号    作者                     日期                    简要描述
 *  1.0         Scott Wang         16-11-29             新增：Create	
 */

namespace app\admin\controller;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\Style\TOC;
use think\Controller;


/**
 * 后台主界面
 *
 * Class main
 *
 * @author      Scott Wang
 * @version     1.0, 16-11-29
 * @since         JSC 1.0
 */
class Main extends Controller
{
    public function main()
    {
        return view();
    }

    /**
     * 生成Word文件
     *
     */
    public function generateWord()
    {
        //获取表单内容
        $title1 = input('input1');
        $body1 = input('textarea1');

        $title2 = input('input2');
        $body2 = input('textarea2');

        $title3 = input('input3');
        $body3 = input('textarea3');

        $title4 = input('input4');
        $body4 = input('textarea4');

        $title5 = input('input5');
        $body5 = input('textarea5');

        $title6 = input('input6');
        $body6 = input('textarea6');

        $title7 = input('input7');
        $body7 = input('textarea7');

        $title8 = input('input8');
        $body8 = input('textarea8');

        $title9 = input('input9');
        $body9 = input('textarea9');

        $phpWord = new PhpWord();

        $section = $phpWord->addSection();

        //预定义样式
        $phpWord->setDefaultFontName("宋体");         //设置默认字体宋体
        $phpWord->setDefaultFontSize(12);               //小四
        $styleTOC = array('tabLeader' => TOC::TAB_LEADER_DOT);
        $phpWord->addTitleStyle(2, array('size' => 16, 'bold' => true, 'name' => '宋体'));
        $phpWord->addTitleStyle(3, array('size' => 12, 'bold' => true, 'name' => '宋体'));

        //添加页脚
        $footer = $section->addFooter();
        $footer->addPreserveText("{PAGE}", array(), array('alignment' => Jc::CENTER));

        //目录
        $section->addText("目录", array(), array('alignment' => Jc::CENTER));
        $section->addTOC(array(), $styleTOC);

        $section->addPageBreak();

        $section->addTitle($title1, 2);
        $section->addText($body1,
            null,
            array('indentation' => array('firstLine' => 480))
        );

        $section->addPageBreak();

        $section->addTitle($title2, 2);
        $section->addText($body2,
            null,
            array('indentation' => array('firstLine' => 480))
        );

        $section->addPageBreak();

        $section->addTitle($title3, 2);
        $section->addText($body3,
            null,
            array('indentation' => array('firstLine' => 480))
        );

        $section->addPageBreak();

        $section->addTitle($title4, 3);
        $section->addText($body4,
            null,
            array('indentation' => array('firstLine' => 480))
        );

        $section->addPageBreak();

        $section->addTitle($title5, 3);
        $section->addText($body5,
            null,
            array('indentation' => array('firstLine' => 480))
        );

        $section->addPageBreak();

        $section->addTitle($title6, 2);
        $section->addText(
            $body6,
            null,
            array('indentation' => array('firstLine' => 480))
        );

        $section->addPageBreak();

        $section->addTitle($title7, 3);
        $section->addText(
            $body7,
            null,
            array('indentation' => array('firstLine' => 480))
        );

        $section->addPageBreak();

        $section->addTitle($title8, 3);
        $section->addText(
            $body8,
            null,
            array('indentation' => array('firstLine' => 480))
        );

        $section->addPageBreak();

        $section->addTitle($title9, 2);
        $section->addText(
            $body9,
            null,
            array('indentation' => array('firstLine' => 480))
        );

        $fileName = "论文小助手Pro" . date("YmdHis") . ".docx";
        $phpWord->save($fileName, "Word2007", true);
    }
}